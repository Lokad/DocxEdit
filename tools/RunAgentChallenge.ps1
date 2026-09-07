param(
    [string] $Challenge,

    [string] $PrivateCase,

    [switch] $List,

    [switch] $DryRun,

    [switch] $PrepareOnly,

    [ValidateSet("read-only", "workspace-write", "danger-full-access")]
    [string] $Sandbox = "workspace-write",

    [string] $Model,

    [int] $TimeoutMinutes = 60,

    [switch] $NoBuild,

    [switch] $NoOutputSchema,

    [switch] $PersistCodexSession,

    [switch] $IgnoreUserConfig,

    [switch] $IgnoreRules,

    [string] $Codex = "codex"
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "DocxCaseCommon.ps1")

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ChallengeRoot = Join-Path $RepoRoot "agent-challenges/challenges"
$SchemaPath = Join-Path $RepoRoot "agent-challenges/final.schema.json"
$PrivateRoot = Join-Path $RepoRoot "private-cases"
$PrivateRunRoot = Join-Path $PrivateRoot "_runs"
$PublicArtifactRoot = Join-Path $RepoRoot "artifacts/agent-challenges"
$CliProject = Join-Path $RepoRoot "src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj"

function ConvertTo-StringList([object] $Value) {
    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [string]) {
        return @([string] $Value)
    }

    $items = @()
    foreach ($item in @($Value)) {
        $text = [string] $item
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $items += $text
        }
    }

    return ,$items
}

function Resolve-ProcessInvocation([string] $FileName, [string[]] $Arguments) {
    $command = Get-Command $FileName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $command) {
        return [pscustomobject]@{
            FileName = $FileName
            Arguments = $Arguments
        }
    }

    $path = [string] $command.Source
    if ([string]::IsNullOrWhiteSpace($path)) {
        $path = [string] $command.Path
    }

    if ([string]::IsNullOrWhiteSpace($path)) {
        return [pscustomobject]@{
            FileName = $FileName
            Arguments = $Arguments
        }
    }

    if ([System.IO.Path]::GetFileName($path).Equals("codex.ps1", [System.StringComparison]::OrdinalIgnoreCase)) {
        $baseDirectory = Split-Path -Parent $path
        $codexJs = Join-Path $baseDirectory "node_modules/@openai/codex/bin/codex.js"
        if (Test-Path -LiteralPath $codexJs) {
            $localNode = Join-Path $baseDirectory "node.exe"
            $nodeCommand = if (Test-Path -LiteralPath $localNode) {
                $localNode
            }
            else {
                $resolvedNode = Get-Command "node" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($null -eq $resolvedNode) {
                    $null
                }
                else {
                    [string] $resolvedNode.Source
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($nodeCommand)) {
                return [pscustomobject]@{
                    FileName = $nodeCommand
                    Arguments = @($codexJs) + $Arguments
                }
            }
        }
    }

    if ([System.IO.Path]::GetExtension($path).Equals(".ps1", [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            FileName = "powershell"
            Arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $path) + $Arguments
        }
    }

    return [pscustomobject]@{
        FileName = $path
        Arguments = $Arguments
    }
}

function Invoke-ProcessCapture([string] $FileName, [string[]] $Arguments, [string] $WorkingDirectory, [string] $StandardInput = $null, [int] $TimeoutSeconds = 0) {
    $invocation = Resolve-ProcessInvocation $FileName $Arguments
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $invocation.FileName
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardInput = $null -ne $StandardInput
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = (($invocation.Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")

    $process = [System.Diagnostics.Process]::Start($startInfo)
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()

    if ($null -ne $StandardInput) {
        $process.StandardInput.Write($StandardInput)
        $process.StandardInput.Close()
    }

    $timedOut = $false
    if ($TimeoutSeconds -gt 0) {
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $timedOut = $true
            try {
                $process.Kill()
            }
            catch {
            }
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

function Get-ChallengeFiles {
    if (-not (Test-Path -LiteralPath $ChallengeRoot)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $ChallengeRoot -Filter "*.json" -File | Sort-Object Name)
}

function Read-ChallengeManifest([string] $ChallengeId) {
    if ([string]::IsNullOrWhiteSpace($ChallengeId)) {
        throw "Pass -Challenge <id>, or use -List to see available challenges."
    }

    if ($ChallengeId -notmatch '^[a-z0-9][a-z0-9-]*$') {
        throw "Challenge id must be kebab-case."
    }

    $path = Join-Path $ChallengeRoot "$ChallengeId.json"
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Unknown challenge '$ChallengeId'. Use -List to see available challenges."
    }

    $manifest = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $id = [string] (Get-ObjectProperty $manifest "id")
    if ($id -ne $ChallengeId) {
        throw "Challenge manifest id must match its filename."
    }

    $title = [string] (Get-ObjectProperty $manifest "title")
    if ([string]::IsNullOrWhiteSpace($title)) {
        throw "Challenge '$ChallengeId' must define a title."
    }

    return $manifest
}

function Resolve-PrivateCase([string] $CaseValue) {
    if (-not (Test-Path -LiteralPath $PrivateRoot)) {
        throw "private-cases/ does not exist. Add a private .docx there or pass -PrivateCase."
    }

    if ([string]::IsNullOrWhiteSpace($CaseValue)) {
        $docxFiles = @(Get-ChildItem -LiteralPath $PrivateRoot -Filter "*.docx" -File)
        if ($docxFiles.Count -ne 1) {
            throw "Expected exactly one private .docx under private-cases/, or pass -PrivateCase."
        }

        $candidate = $docxFiles[0].FullName
    }
    elseif ([System.IO.Path]::IsPathRooted($CaseValue)) {
        $candidate = $CaseValue
    }
    else {
        $privateCandidate = Join-Path $PrivateRoot $CaseValue
        if (Test-Path -LiteralPath $privateCandidate) {
            $candidate = $privateCandidate
        }
        else {
            $candidate = Join-Path $RepoRoot $CaseValue
        }
    }

    $casePath = (Resolve-Path -LiteralPath $candidate).Path
    Assert-PrivatePath $casePath "Private case"

    if (-not [System.IO.Path]::GetExtension($casePath).Equals(".docx", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Private case must be a .docx file."
    }

    return $casePath
}

function New-DocxEditWrapper([string] $RunDirectory) {
    $wrapperPath = Join-Path $RunDirectory "docxedit.ps1"
    $content = @'
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $Arguments
)

& dotnet run --no-build --project "__CLI_PROJECT__" -- @Arguments
exit $LASTEXITCODE
'@

    $content = $content.Replace("__CLI_PROJECT__", $CliProject)
    Set-Content -LiteralPath $wrapperPath -Value $content -Encoding UTF8
    return $wrapperPath
}

function New-ChallengePrompt([object] $Manifest, [string] $RunDirectory, [string] $InputPath, [bool] $UseOutputSchema) {
    $id = [string] (Get-ObjectProperty $Manifest "id")
    $title = [string] (Get-ObjectProperty $Manifest "title")
    $requiresOutput = [bool] (Get-ObjectProperty $Manifest "requiresOutputDocx")
    $task = ConvertTo-StringList (Get-ObjectProperty $Manifest "task")
    $successCriteria = ConvertTo-StringList (Get-ObjectProperty $Manifest "successCriteria")

    $lines = [System.Collections.Generic.List[string]]::new()
    [void] $lines.Add("# DocxEdit Agent Challenge: $title")
    [void] $lines.Add("")
    [void] $lines.Add('You are a fresh Codex agent evaluating whether `docxedit` is understandable enough for private .docx editing tasks.')
    [void] $lines.Add("")
    [void] $lines.Add("Workspace:")
    [void] $lines.Add('- Run directory: `' + $RunDirectory + '`')
    [void] $lines.Add('- Private input copy: `.\input.docx`')
    [void] $lines.Add('- DocxEdit wrapper: `.\docxedit.ps1`')
    [void] $lines.Add("- Put all generated patches, logs, and edited documents in this run directory.")
    [void] $lines.Add("")
    [void] $lines.Add("Rules:")
    [void] $lines.Add('- Use only `docxedit` for .docx inspection and editing.')
    [void] $lines.Add("- You may use shell commands to create patch files, list files, and run validation commands.")
    [void] $lines.Add('- Do not unzip the .docx, inspect `word/*.xml`, use Word COM, use Office automation, use Open XML SDK, or use python-docx.')
    [void] $lines.Add('- Do not modify `.\input.docx` in place; write a separate output document when an output is required.')
    [void] $lines.Add("- Do not quote private document text in your final response. Use IDs, counts, diagnostics, text lengths, and short operation descriptions.")
    [void] $lines.Add("- If a command prints private text while you are exploring, do not repeat that text in the final response.")
    [void] $lines.Add("")
    [void] $lines.Add("Useful commands:")
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 help patch`')
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 changes .\input.docx --json`')
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 read .\input.docx --headers-footers --view markup --max-text 200`')
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 dump .\input.docx --id M.P0001 --runs --view markup --json --max-text 200`')
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 check .\input.docx .\edits.docxpatch --json --report .\check-report.json`')
    [void] $lines.Add('- `powershell -NoProfile -ExecutionPolicy Bypass -File .\docxedit.ps1 apply .\input.docx .\edits.docxpatch --output .\edited.docx --json --report .\apply-report.json`')
    [void] $lines.Add("")
    [void] $lines.Add("Task:")
    foreach ($item in $task) {
        [void] $lines.Add("- $item")
    }
    [void] $lines.Add("")
    [void] $lines.Add("Success criteria:")
    foreach ($item in $successCriteria) {
        [void] $lines.Add("- $item")
    }
    [void] $lines.Add("")
    [void] $lines.Add("Output document required: $requiresOutput")
    if ($UseOutputSchema) {
        [void] $lines.Add("Final response must match the provided JSON schema. Keep all fields private-text-free.")
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-CodexCommandArgs([string] $RunDirectory, [string] $FinalPath, [bool] $UseOutputSchema) {
    $args = @(
        "exec",
        "--json",
        "--color",
        "never",
        "--cd",
        $RunDirectory,
        "--sandbox",
        $Sandbox
    )

    if (-not $PersistCodexSession) {
        $args += "--ephemeral"
    }

    if ($IgnoreUserConfig) {
        $args += "--ignore-user-config"
    }

    if ($IgnoreRules) {
        $args += "--ignore-rules"
    }

    if (-not [string]::IsNullOrWhiteSpace($Model)) {
        $args += @("--model", $Model)
    }

    if ($UseOutputSchema) {
        $args += @("--output-schema", $SchemaPath)
    }

    $args += @("-o", $FinalPath, "-")
    return ,$args
}

function Read-CodexEvents([string] $EventsPath) {
    $commandsById = @{}
    $threadId = $null
    $usage = $null
    $parseErrors = 0

    if (-not (Test-Path -LiteralPath $EventsPath)) {
        return [pscustomobject]@{
            ThreadId = $null
            Usage = $null
            Commands = @()
            ParseErrors = 0
        }
    }

    foreach ($line in Get-Content -LiteralPath $EventsPath) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $event = $line | ConvertFrom-Json
        }
        catch {
            $parseErrors++
            continue
        }

        $type = [string] (Get-ObjectProperty $event "type")
        if ($type -eq "thread.started") {
            $threadId = [string] (Get-ObjectProperty $event "thread_id")
            continue
        }

        if ($type -eq "turn.completed") {
            $usage = Get-ObjectProperty $event "usage"
            continue
        }

        if ($type -ne "item.started" -and $type -ne "item.completed") {
            continue
        }

        $item = Get-ObjectProperty $event "item"
        if ($null -eq $item) {
            continue
        }

        $itemType = [string] (Get-ObjectProperty $item "type")
        if ($itemType -ne "command_execution") {
            continue
        }

        $itemId = [string] (Get-ObjectProperty $item "id")
        if ([string]::IsNullOrWhiteSpace($itemId)) {
            $itemId = "command-$($commandsById.Count + 1)"
        }

        $commandText = [string] (Get-ObjectProperty $item "command")
        $status = [string] (Get-ObjectProperty $item "status")
        $exitCode = Get-ObjectProperty $item "exit_code"
        $commandsById[$itemId] = [pscustomobject]@{
            Id = $itemId
            Command = $commandText
            Status = $status
            ExitCode = $exitCode
        }
    }

    return [pscustomobject]@{
        ThreadId = $threadId
        Usage = $usage
        Commands = @($commandsById.Values | Sort-Object Id)
        ParseErrors = $parseErrors
    }
}

function Test-AnyCommand([object[]] $Commands, [string] $Pattern) {
    foreach ($command in @($Commands)) {
        $text = [string] $command.Command
        if ($text -match $Pattern) {
            return $true
        }
    }

    return $false
}

function Invoke-PostChecks([string] $RunDirectory) {
    $postRoot = Join-Path $RunDirectory "postcheck"
    New-Item -ItemType Directory -Force -Path $postRoot | Out-Null

    $results = @()
    $docxFiles = @(Get-ChildItem -LiteralPath $RunDirectory -Filter "*.docx" -File | Where-Object { $_.Name -ne "input.docx" } | Sort-Object Name)
    foreach ($docx in $docxFiles) {
        $safeName = [System.IO.Path]::GetFileNameWithoutExtension($docx.Name)
        $read = Invoke-ProcessCapture "dotnet" @(
            "run",
            "--no-build",
            "--project",
            $CliProject,
            "--",
            "read",
            $docx.FullName,
            "--max-text",
            "0",
            "--json") $RepoRoot
        $changes = Invoke-ProcessCapture "dotnet" @(
            "run",
            "--no-build",
            "--project",
            $CliProject,
            "--",
            "changes",
            $docx.FullName,
            "--json") $RepoRoot

        Set-Content -LiteralPath (Join-Path $postRoot "$safeName.read.json") -Value $read.StdOut -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $postRoot "$safeName.read.stderr.log") -Value $read.StdErr -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $postRoot "$safeName.changes.json") -Value $changes.StdOut -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $postRoot "$safeName.changes.stderr.log") -Value $changes.StdErr -Encoding UTF8

        $results += [pscustomobject]@{
            File = $docx.Name
            Bytes = $docx.Length
            ReadExitCode = $read.ExitCode
            ChangesExitCode = $changes.ExitCode
        }
    }

    return ,$results
}

function Write-JsonFile([string] $Path, [object] $Value) {
    $Value | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
}

if ($List -or [string]::IsNullOrWhiteSpace($Challenge)) {
    $files = Get-ChallengeFiles
    if ($files.Count -eq 0) {
        Write-Host "no agent challenges found"
        exit 0
    }

    foreach ($file in $files) {
        $manifest = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
        $id = [string] (Get-ObjectProperty $manifest "id")
        $title = [string] (Get-ObjectProperty $manifest "title")
        $requiresOutput = [bool] (Get-ObjectProperty $manifest "requiresOutputDocx")
        Write-Host "$id`t$title`toutput-docx=$requiresOutput"
    }

    if (-not $List) {
        Write-Host ""
        Write-Host "pass -Challenge <id> to run or -DryRun to preview the prompt"
    }

    exit 0
}

$manifest = Read-ChallengeManifest $Challenge
$useOutputSchema = -not $NoOutputSchema

if ($DryRun) {
    $prompt = New-ChallengePrompt $manifest "<run-directory>" ".\input.docx" $useOutputSchema
    Write-Host $prompt
    Write-Host ""
    Write-Host "codex command:"
    $previewArgs = Get-CodexCommandArgs "<run-directory>" "<final-response>" $useOutputSchema
    Write-Host "$Codex $($previewArgs -join ' ')"
    exit 0
}

$privatePath = Resolve-PrivateCase $PrivateCase
$runId = [DateTimeOffset]::UtcNow.ToString("yyyyMMddTHHmmssZ")
$challengeId = [string] (Get-ObjectProperty $manifest "id")
$runDirectory = Join-Path $PrivateRunRoot (Join-Path $challengeId $runId)
New-Item -ItemType Directory -Force -Path $runDirectory | Out-Null
$runDirectory = (Resolve-Path -LiteralPath $runDirectory).Path
Assert-PrivatePath $runDirectory "Run directory"

$inputCopy = Join-Path $runDirectory "input.docx"
Copy-Item -LiteralPath $privatePath -Destination $inputCopy
$inputHashBefore = Get-FileHashSha256 $inputCopy
$privateHash = Get-FileHashSha256 $privatePath

$wrapperPath = New-DocxEditWrapper $runDirectory
$prompt = New-ChallengePrompt $manifest $runDirectory $inputCopy $useOutputSchema
$promptPath = Join-Path $runDirectory "prompt.md"
Set-Content -LiteralPath $promptPath -Value $prompt -Encoding UTF8

if (-not $NoBuild) {
    $build = Invoke-ProcessCapture "dotnet" @("build", $CliProject, "--nologo", "--verbosity", "minimal") $RepoRoot
    Set-Content -LiteralPath (Join-Path $runDirectory "build.stdout.log") -Value $build.StdOut -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $runDirectory "build.stderr.log") -Value $build.StdErr -Encoding UTF8
    if ($build.ExitCode -ne 0) {
        throw "CLI build failed. See ignored run artifact logs."
    }
}

if ($PrepareOnly) {
    $prepareSummaryPath = Join-Path $runDirectory "prepare-summary.json"
    $prepareSummary = [pscustomobject]@{
        ChallengeId = $challengeId
        Title = [string] (Get-ObjectProperty $manifest "title")
        RunId = $runId
        Prepared = $true
        PrivateInputHashSha256 = $privateHash
        InputCopyHashSha256 = $inputHashBefore
        ArtifactPaths = [pscustomobject]@{
            RunDirectory = Get-RepoRelativePath $runDirectory
            Prompt = Get-RepoRelativePath $promptPath
            Wrapper = Get-RepoRelativePath $wrapperPath
            InputCopy = Get-RepoRelativePath $inputCopy
        }
    }

    Write-JsonFile $prepareSummaryPath $prepareSummary
    Write-Host "agent challenge '$challengeId': PREPARED"
    Write-Host "artifact: $(Get-RepoRelativePath $prepareSummaryPath)"
    Write-Host "run-directory: $(Get-RepoRelativePath $runDirectory)"
    exit 0
}

$finalPath = Join-Path $runDirectory $(if ($useOutputSchema) { "final-response.json" } else { "final-response.txt" })
$eventsPath = Join-Path $runDirectory "events.jsonl"
$stderrPath = Join-Path $runDirectory "codex.stderr.log"
$codexArgs = Get-CodexCommandArgs $runDirectory $finalPath $useOutputSchema
$codexVersion = Invoke-ProcessCapture $Codex @("--version") $RepoRoot
$trackedStatusBefore = @(& git -C $RepoRoot status --short --untracked-files=no)

$run = Invoke-ProcessCapture $Codex $codexArgs $runDirectory $prompt ($TimeoutMinutes * 60)
Set-Content -LiteralPath $eventsPath -Value $run.StdOut -Encoding UTF8
Set-Content -LiteralPath $stderrPath -Value $run.StdErr -Encoding UTF8

$trackedStatusAfter = @(& git -C $RepoRoot status --short --untracked-files=no)
$inputHashAfter = Get-FileHashSha256 $inputCopy
$events = Read-CodexEvents $eventsPath
$postChecks = Invoke-PostChecks $runDirectory

$commands = @($events.Commands)
$usedForbiddenPattern = '(?i)(Expand-Archive|unzip|python-docx|DocumentFormat\.OpenXml|Word\.Application|winword|word[\\/][^ ]*\.xml|System\.IO\.Compression\.ZipFile|ZipArchive)'
$metrics = [pscustomobject]@{
    CommandCount = $commands.Count
    UsedDocxEdit = Test-AnyCommand $commands '(?i)(docxedit\.ps1|Lokad\.DocxEdit\.Cli|docxedit)'
    UsedChanges = Test-AnyCommand $commands '(?i)\bchanges\b'
    UsedMarkupView = Test-AnyCommand $commands '(?i)--view\s+markup'
    UsedCheck = Test-AnyCommand $commands '(?i)\bcheck\b'
    UsedApply = Test-AnyCommand $commands '(?i)\bapply\b'
    UsedForbiddenDocxInspection = Test-AnyCommand $commands $usedForbiddenPattern
}

$outputDocx = @(Get-ChildItem -LiteralPath $runDirectory -Filter "*.docx" -File | Where-Object { $_.Name -ne "input.docx" } | ForEach-Object { $_.Name } | Sort-Object)
$summary = [pscustomobject]@{
    ChallengeId = $challengeId
    Title = [string] (Get-ObjectProperty $manifest "title")
    RunId = $runId
    Success = ($run.ExitCode -eq 0 -and -not $run.TimedOut)
    ExitCode = $run.ExitCode
    TimedOut = $run.TimedOut
    Sandbox = $Sandbox
    Model = if ([string]::IsNullOrWhiteSpace($Model)) { $null } else { $Model }
    Ephemeral = -not $PersistCodexSession
    UsedOutputSchema = $useOutputSchema
    CodexVersion = ($codexVersion.StdOut + $codexVersion.StdErr).Trim()
    ThreadId = $events.ThreadId
    Usage = $events.Usage
    Metrics = $metrics
    OutputDocx = $outputDocx
    PostChecks = $postChecks
    InputCopyModified = ($inputHashBefore -ne $inputHashAfter)
    PrivateInputHashSha256 = $privateHash
    InputCopyHashBeforeSha256 = $inputHashBefore
    InputCopyHashAfterSha256 = $inputHashAfter
    TrackedStatusBefore = $trackedStatusBefore
    TrackedStatusAfter = $trackedStatusAfter
    EventParseErrors = $events.ParseErrors
    Commands = $commands
    ArtifactPaths = [pscustomobject]@{
        RunDirectory = Get-RepoRelativePath $runDirectory
        Prompt = Get-RepoRelativePath $promptPath
        Events = Get-RepoRelativePath $eventsPath
        Stderr = Get-RepoRelativePath $stderrPath
        FinalResponse = Get-RepoRelativePath $finalPath
    }
}

$summaryPath = Join-Path $runDirectory "summary.json"
Write-JsonFile $summaryPath $summary
$publicRunDirectory = Join-Path $PublicArtifactRoot (Join-Path $challengeId $runId)
New-Item -ItemType Directory -Force -Path $publicRunDirectory | Out-Null
$publicSummaryPath = Join-Path $publicRunDirectory "summary.json"
Write-JsonFile $publicSummaryPath (New-SanitizedChallengeSummary $summary)

Write-Host "agent challenge '$challengeId': $(if ($summary.Success) { 'OK' } else { 'FAILED' })"
Write-Host "artifact: $(Get-RepoRelativePath $summaryPath)"
Write-Host "sanitized-artifact: $(Get-RepoRelativePath $publicSummaryPath)"
Write-Host "codex: exit=$($run.ExitCode) timedOut=$($run.TimedOut) thread=$($events.ThreadId)"
Write-Host "commands: count=$($metrics.CommandCount) docxedit=$($metrics.UsedDocxEdit) changes=$($metrics.UsedChanges) check=$($metrics.UsedCheck) apply=$($metrics.UsedApply) forbidden-docx-inspection=$($metrics.UsedForbiddenDocxInspection)"
Write-Host "output-docx: $($outputDocx.Count)"
Write-Host "input-copy-modified: $($summary.InputCopyModified)"

if ($run.ExitCode -ne 0 -or $run.TimedOut) {
    exit 1
}
