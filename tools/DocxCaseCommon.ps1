# Shared helpers for the DocxEdit PowerShell tooling.
# Dot-sourced by CheckPrivateCase.ps1 and RunAgentChallenge.ps1.
# Callers must define $RepoRoot (process working directory, git checks, relative paths);
# Assert-PrivatePath additionally needs $PrivateRoot.
# RunAgentChallenge.ps1 adapts process invocation (ps1 and codex-shim resolution) and forwards to the shared Invoke-ProcessCapture core below.

function Test-IsUnderPath([string] $Path, [string] $Root) {
    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $resolvedRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $lexicallyInside = $resolvedPath.Equals($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $resolvedPath.StartsWith($resolvedRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $lexicallyInside) {
        return $false
    }

    return -not (Test-PathHasReparsePoint $resolvedPath)
}

function Test-PathHasReparsePoint([string] $FullPath) {
    $cursor = $FullPath
    while ($null -ne $cursor -and $cursor -ne "") {
        if (Test-Path -LiteralPath $cursor) {
            try {
                $attributes = (Get-Item -LiteralPath $cursor -Force).Attributes
                if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    return $true
                }
            } catch {
            }
            if ($cursor -eq [System.IO.Path]::GetPathRoot($cursor)) {
                break
            }
            $cursor = [System.IO.Path]::GetDirectoryName($cursor)
            continue
        }
        $cursor = [System.IO.Path]::GetDirectoryName($cursor)
    }
    return $false
}

function Get-RepoRelativePath([string] $Path) {
    $root = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $rootUri = [System.Uri]::new($root)
    $pathUri = [System.Uri]::new($fullPath)
    return [System.Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('\', '/')
}

function Assert-PrivatePath([string] $Path, [string] $Kind) {
    if (-not (Test-IsUnderPath $Path $PrivateRoot)) {
        throw "$Kind must live under private-cases/."
    }

    $relative = Get-RepoRelativePath $Path
    $tracked = & git -C $RepoRoot ls-files -- $relative
    if ($tracked) {
        throw "$Kind must not be tracked by git."
    }

    & git -C $RepoRoot check-ignore -q -- $relative
    if ($LASTEXITCODE -ne 0) {
        throw "$Kind must be ignored by git."
    }
}

function ConvertTo-ProcessArgument([string] $Argument) {
    if ([string]::IsNullOrEmpty($Argument)) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    # CommandLineToArgvW rules: backslashes are literal unless followed by a
    # quote, so double every backslash run that precedes a quote or the end.
    $escaped = $Argument -replace '(\\*)"', '$1$1\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Get-CaptureRemainingMs([DateTime] $Deadline) {
    if ($Deadline -eq [DateTime]::MaxValue) {
        return 2147483647
    }

    $remaining = [int](($Deadline - [DateTime]::UtcNow).TotalMilliseconds)
    if ($remaining -lt 0) {
        return 0
    }

    return $remaining
}

function Invoke-ProcessCapture([string] $FileName, [string[]] $Arguments, [int] $TimeoutSeconds = 0, [string] $WorkingDirectory = "", [string] $StandardInput = $null) {
    $directory = $WorkingDirectory
    if ([string]::IsNullOrWhiteSpace($directory)) {
        $directory = $RepoRoot
    }

    $deadline = [DateTime]::MaxValue
    if ($TimeoutSeconds -gt 0) {
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    }

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FileName
    $startInfo.WorkingDirectory = $directory
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardInput = $null -ne $StandardInput
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")

    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        # Drain both streams concurrently: sequential ReadToEnd calls deadlock
        # once the child fills the pipe nobody is reading.
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $timedOut = $false
        $cleanupFailed = $false

        if ($null -ne $StandardInput) {
            $stdinTask = $process.StandardInput.WriteAsync($StandardInput)
            if (-not $stdinTask.Wait((Get-CaptureRemainingMs $deadline))) {
                $timedOut = $true
            }

            try { $process.StandardInput.Close() } catch { }
        }

        if (-not $timedOut) {
            if ($TimeoutSeconds -gt 0) {
                if (-not $process.WaitForExit((Get-CaptureRemainingMs $deadline))) {
                    $timedOut = $true
                }
            }
            else {
                $process.WaitForExit()
            }
        }

        if ($timedOut) {
            Stop-ProcessTree $process.Id
            try { $process.Kill() } catch { }
            try {
                if (-not $process.WaitForExit(5000)) {
                    $cleanupFailed = $true
                }
            } catch {
                $cleanupFailed = $true
            }
        }

        $drained = $false
        if ($TimeoutSeconds -gt 0) {
            $drainBudget = Get-CaptureRemainingMs $deadline
            if ($drainBudget -gt 0 -and [System.Threading.Tasks.Task]::WaitAll(@($stdoutTask, $stderrTask), $drainBudget)) {
                $drained = $true
            }

            if (-not $drained) {
                $timedOut = $true
                Stop-ProcessTree $process.Id
                try {
                    if (-not ([System.Threading.Tasks.Task]::WaitAll(@($stdoutTask, $stderrTask), 5000))) {
                        $cleanupFailed = $true
                    }
                    else {
                        $drained = $true
                    }
                } catch {
                    $cleanupFailed = $true
                }
            }
        }
        else {
            $stdoutTask.Wait(-1)
            $stderrTask.Wait(-1)
            $drained = $true
        }

        [pscustomobject]@{
            ExitCode = if ($timedOut) { -1 } else { $process.ExitCode }
            TimedOut = $timedOut
            CleanupFailed = $cleanupFailed
            StdOut = if ($drained) { $stdoutTask.Result } else { "" }
            StdErr = if ($drained) { $stderrTask.Result } else { "" }
        }
    }
    finally {
        $process.Dispose()
    }
}

function Stop-ProcessTree([int] $ProcessId) {
    try {
        $children = @(Get-CimInstance Win32_Process -Filter ("ParentProcessId=" + $ProcessId) -ErrorAction SilentlyContinue)
    } catch {
        return
    }

    foreach ($child in $children) {
        Stop-ProcessTree $child.ProcessId
        try { Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue } catch { }
    }
}
function Get-FileHashSha256([string] $Path) {
    if ($null -eq (Get-Command Get-FileHash -ErrorAction SilentlyContinue)) {
        throw "File hashing needs the Get-FileHash cmdlet (pwsh 7+ or a full Windows PowerShell 5.1 install); refusing to continue with an unverified copy."
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Get-ObjectProperty([object] $Object, [string] $Name) {
    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}


function Get-StableCaseKey([string] $CaseId) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($CaseId.ToLowerInvariant())
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant().Substring(0, 16)
}

function New-SanitizedChallengeSummary([object] $Summary) {
    $metrics = Get-ObjectProperty $Summary "Metrics"
    $postChecks = @()
    foreach ($postCheck in @(Get-ObjectProperty $Summary "PostChecks")) {
        $postChecks += [pscustomobject]@{
            Bytes = $postCheck.Bytes
            ReadExitCode = $postCheck.ReadExitCode
            ChangesExitCode = $postCheck.ChangesExitCode
        }
    }

    return [pscustomobject]@{
        ChallengeId = [string](Get-ObjectProperty $Summary "ChallengeId")
        Title = [string](Get-ObjectProperty $Summary "Title")
        RunId = [string](Get-ObjectProperty $Summary "RunId")
        Success = [bool](Get-ObjectProperty $Summary "Success")
        ExitCode = Get-ObjectProperty $Summary "ExitCode"
        TimedOut = [bool](Get-ObjectProperty $Summary "TimedOut")
        Sandbox = [string](Get-ObjectProperty $Summary "Sandbox")
        Model = [string](Get-ObjectProperty $Summary "Model")
        Ephemeral = [bool](Get-ObjectProperty $Summary "Ephemeral")
        UsedOutputSchema = [bool](Get-ObjectProperty $Summary "UsedOutputSchema")
        CodexVersion = [string](Get-ObjectProperty $Summary "CodexVersion")
        ThreadId = [string](Get-ObjectProperty $Summary "ThreadId")
        Usage = Get-ObjectProperty $Summary "Usage"
        EventParseErrors = Get-ObjectProperty $Summary "EventParseErrors"
        InputCopyModified = [bool](Get-ObjectProperty $Summary "InputCopyModified")
        PrivateInputHashSha256 = [string](Get-ObjectProperty $Summary "PrivateInputHashSha256")
        InputCopyHashBeforeSha256 = [string](Get-ObjectProperty $Summary "InputCopyHashBeforeSha256")
        InputCopyHashAfterSha256 = [string](Get-ObjectProperty $Summary "InputCopyHashAfterSha256")
        CommandCount = [int](Get-ObjectProperty $metrics "CommandCount")
        UsedDocxEdit = [bool](Get-ObjectProperty $metrics "UsedDocxEdit")
        UsedChanges = [bool](Get-ObjectProperty $metrics "UsedChanges")
        UsedMarkupView = [bool](Get-ObjectProperty $metrics "UsedMarkupView")
        UsedCheck = [bool](Get-ObjectProperty $metrics "UsedCheck")
        UsedApply = [bool](Get-ObjectProperty $metrics "UsedApply")
        UsedForbiddenDocxInspection = [bool](Get-ObjectProperty $metrics "UsedForbiddenDocxInspection")
        TaskCompleted = Get-ObjectProperty $Summary "TaskCompleted"
        Verdict = [string](Get-ObjectProperty $Summary "Verdict")
        FailureReasons = @((Get-ObjectProperty $Summary "FailureReasons") | ForEach-Object { [string]$_ })
        OutputDocxCount = @((Get-ObjectProperty $Summary "OutputDocx")).Count
        PostChecks = $postChecks
    }
}

function New-SanitizedPrivateCaseSummary([object] $Summary, [string] $CaseKey) {
    $diagnosticCodes = @((Get-ObjectProperty $Summary "Diagnostics") | ForEach-Object { [string](Get-ObjectProperty $_ "Code") } | Sort-Object -Unique)
    $changeTypes = @{}
    $byType = Get-ObjectProperty (Get-ObjectProperty $Summary "Changes") "ByType"
    if ($null -ne $byType) {
        foreach ($key in @($byType.Keys)) {
            $changeTypes[[string] $key] = [int] $byType[$key]
        }
    }

    return [pscustomobject]@{
        CaseKey = $CaseKey
        RunId = [string](Get-ObjectProperty $Summary "RunId")
        ValidateOnly = [bool](Get-ObjectProperty $Summary "ValidateOnly")
        Success = [bool](Get-ObjectProperty $Summary "Success")
        InputBytes = [long](Get-ObjectProperty $Summary "InputBytes")
        Paragraphs = [long](Get-ObjectProperty (Get-ObjectProperty $Summary "Structure") "Paragraphs")
        Tables = [long](Get-ObjectProperty (Get-ObjectProperty $Summary "Structure") "Tables")
        Images = [long](Get-ObjectProperty (Get-ObjectProperty $Summary "Structure") "ScannerVisibleImages")
        Sections = [long](Get-ObjectProperty (Get-ObjectProperty $Summary "Structure") "Sections")
        Styles = [long](Get-ObjectProperty (Get-ObjectProperty $Summary "Structure") "Styles")
        ChangeTotal = [int](Get-ObjectProperty (Get-ObjectProperty $Summary "Changes") "Total")
        ChangeByType = $changeTypes
        DiagnosticCodes = $diagnosticCodes
        DiagnosticCount = @((Get-ObjectProperty $Summary "Diagnostics")).Count
        AggregateFailures = @((Get-ObjectProperty $Summary "AggregateFailures"))
    }
}

function Read-FinalResponse([string] $FinalPath, [bool] $UseOutputSchema) {
    if (-not $UseOutputSchema) {
        return [pscustomobject]@{ Final = $null; ParseError = "unstructured-final" }
    }

    if (-not (Test-Path -LiteralPath $FinalPath)) {
        return [pscustomobject]@{ Final = $null; ParseError = "final-missing" }
    }

    try {
        $parsed = Get-Content -LiteralPath $FinalPath -Raw | ConvertFrom-Json
    } catch {
        return [pscustomobject]@{ Final = $null; ParseError = "final-unparseable" }
    }

    if ((Get-ObjectProperty $parsed "completed") -isnot [bool]) {
        return [pscustomobject]@{ Final = $null; ParseError = "final-invalid" }
    }

    return [pscustomobject]@{ Final = $parsed; ParseError = $null }
}

function Get-ChallengeOutcome([bool] $ProcessOk, [object] $Final, [string] $FinalParseError, [bool] $RequiresOutput, [string] $Applicability, [string[]] $OutputDocx, [object[]] $PostChecks, [bool] $InputModified, [bool] $ForbiddenInspection, [int] $CommandCount, [bool] $UsedDocxEdit) {
    $reasons = [System.Collections.Generic.List[string]]::new()
    if (-not $ProcessOk) {
        $reasons.Add("process-failed")
    }

    $taskCompleted = $null
    $declaredOutput = $null
    if ($null -ne $Final) {
        $taskCompleted = [bool](Get-ObjectProperty $Final "completed")
        $declaredOutput = Get-ObjectProperty $Final "output_docx"
        if (-not $taskCompleted) {
            $reasons.Add("final-incomplete")
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($FinalParseError)) {
        $reasons.Add($FinalParseError)
    }

    $outputNames = @($OutputDocx)
    if ($RequiresOutput) {
        if ($outputNames.Count -eq 0) {
            $reasons.Add("no-output")
        }
        elseif ($null -ne $declaredOutput -and -not ($outputNames -contains $declaredOutput)) {
            $reasons.Add("output-mismatch")
        }
    }

    foreach ($post in @($PostChecks)) {
        if ([int](Get-ObjectProperty $post "ReadExitCode") -ne 0 -or [int](Get-ObjectProperty $post "ChangesExitCode") -ne 0) {
            $reasons.Add("postcheck-failed")
            break
        }
    }

    if ($InputModified) {
        $reasons.Add("input-mutated")
    }

    if ($ForbiddenInspection) {
        $reasons.Add("forbidden-inspection")
    }

    $distinct = @($reasons | Sort-Object -Unique)
    if ($null -eq $Final -and $ProcessOk) {
        return [pscustomobject]@{
            Verdict = "unevaluated"
            TaskCompleted = $null
            FailureReasons = $distinct
        }
    }

    if ($distinct.Count -eq 0) {
        return [pscustomobject]@{
            Verdict = "passed"
            TaskCompleted = $taskCompleted
            FailureReasons = @()
        }
    }

    if ($taskCompleted -eq $false -and $outputNames.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($Applicability) -and $UsedDocxEdit -and $CommandCount -gt 0) {
        $benign = @($distinct | Where-Object { $_ -eq "final-incomplete" -or $_ -eq "no-output" })
        if ($benign.Count -eq $distinct.Count) {
            return [pscustomobject]@{
                Verdict = "not-applicable"
                TaskCompleted = $false
                FailureReasons = @("feature-absent")
            }
        }
    }

    return [pscustomobject]@{
        Verdict = "failed"
        TaskCompleted = $taskCompleted
        FailureReasons = $distinct
    }
}
