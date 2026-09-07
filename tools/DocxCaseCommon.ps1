# Shared helpers for the DocxEdit PowerShell tooling.
# Dot-sourced by CheckPrivateCase.ps1 and RunAgentChallenge.ps1.
# Callers must define $RepoRoot (process working directory, git checks, relative paths);
# Assert-PrivatePath additionally needs $PrivateRoot.
# RunAgentChallenge.ps1 keeps its own Invoke-ProcessCapture (timeouts and stdin).

function Test-IsUnderPath([string] $Path, [string] $Root) {
    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $resolvedRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $resolvedPath.Equals($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $resolvedPath.StartsWith($resolvedRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
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

function Invoke-ProcessCapture([string] $FileName, [string[]] $Arguments, [int] $TimeoutSeconds = 0) {
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FileName
    $startInfo.WorkingDirectory = $RepoRoot
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")

    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        # Drain both streams concurrently: sequential ReadToEnd calls deadlock
        # once the child fills the pipe nobody is reading.
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()

        $timedOut = $false
        if ($TimeoutSeconds -gt 0) {
            if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
                $timedOut = $true
                try { $process.Kill() } catch { }
                $process.WaitForExit()
            }
        }
        else {
            $process.WaitForExit()
        }

        [pscustomobject]@{
            ExitCode = if ($timedOut) { -1 } else { $process.ExitCode }
            TimedOut = $timedOut
            StdOut = $stdoutTask.GetAwaiter().GetResult()
            StdErr = $stderrTask.GetAwaiter().GetResult()
        }
    }
    finally {
        $process.Dispose()
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
