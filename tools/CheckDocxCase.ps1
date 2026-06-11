param(
    [Parameter(Mandatory = $true)]
    [string] $Case
)

$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$CaseRoot = Join-Path $RepoRoot "edit-cases/cases"
$ArtifactRoot = Join-Path $RepoRoot "artifacts/edit-cases"
$CliProject = Join-Path $RepoRoot "src/DocxEdit.Cli/DocxEdit.Cli.csproj"

function ConvertTo-ProcessArgument([string] $Argument) {
    if ([string]::IsNullOrEmpty($Argument)) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    return '"' + $Argument.Replace('"', '\"') + '"'
}

function Invoke-ProcessCapture([string] $FileName, [string[]] $Arguments) {
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FileName
    $startInfo.WorkingDirectory = $RepoRoot
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")

    $process = [System.Diagnostics.Process]::Start($startInfo)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    [pscustomobject]@{
        ExitCode = $process.ExitCode
        StdOut = $stdout
        StdErr = $stderr
    }
}

function Resolve-CasePath([string] $CaseValue) {
    if ([System.IO.Path]::IsPathRooted($CaseValue) -or (Test-Path -LiteralPath $CaseValue)) {
        return (Resolve-Path -LiteralPath $CaseValue).Path
    }

    $name = if ($CaseValue.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) { $CaseValue } else { "$CaseValue.json" }
    return (Resolve-Path -LiteralPath (Join-Path $CaseRoot $name)).Path
}

function Get-ManifestText($Value) {
    if ($Value -is [System.Array]) {
        return ($Value -join [Environment]::NewLine)
    }

    return [string] $Value
}

function Add-ZipEntry([System.IO.Compression.ZipArchive] $Archive, [string] $Name, [string] $Text) {
    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $stream.Write($bytes, 0, $bytes.Length)
    }
    finally {
        $stream.Dispose()
    }
}

function New-DocxFromBody([string] $Path, [string] $BodyXml) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path
    }

    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipEntry $archive "[Content_Types].xml" @"
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@
        Add-ZipEntry $archive "_rels/.rels" @"
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@
        Add-ZipEntry $archive "word/_rels/document.xml.rels" @"
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
"@
        Add-ZipEntry $archive "word/document.xml" @"
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
$BodyXml
  </w:body>
</w:document>
"@
    }
    finally {
        $archive.Dispose()
    }
}

function Assert-StringArrayEquals([string[]] $Expected, [string[]] $Actual, [string] $Label) {
    if ($Expected.Count -ne $Actual.Count) {
        throw "$Label count mismatch. Expected $($Expected.Count), found $($Actual.Count)."
    }

    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if ($Expected[$i] -ne $Actual[$i]) {
            throw "$Label mismatch at index $i. Expected '$($Expected[$i])', found '$($Actual[$i])'."
        }
    }
}

$casePath = Resolve-CasePath $Case
$manifest = Get-Content -LiteralPath $casePath -Raw | ConvertFrom-Json
$caseId = [string] $manifest.id
if ([string]::IsNullOrWhiteSpace($caseId) -or $caseId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
    throw "Case manifest must define a filename-safe id."
}

$bodyXml = Get-ManifestText $manifest.input.bodyXml
$patchText = Get-ManifestText $manifest.patch
if ([string]::IsNullOrWhiteSpace($bodyXml) -or [string]::IsNullOrWhiteSpace($patchText)) {
    throw "Case '$caseId' must define input.bodyXml and patch."
}

$runId = [DateTimeOffset]::UtcNow.ToString("yyyyMMddTHHmmssZ")
$artifactDir = Join-Path $ArtifactRoot (Join-Path $caseId $runId)
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null

$inputPath = Join-Path $artifactDir "input.docx"
$patchPath = Join-Path $artifactDir "edit.docxpatch"
$outputPath = Join-Path $artifactDir "output.docx"
$applyReportPath = Join-Path $artifactDir "apply-report.json"
$summaryPath = Join-Path $artifactDir "summary.json"

New-DocxFromBody $inputPath $bodyXml
Set-Content -LiteralPath $patchPath -Value $patchText

$build = Invoke-ProcessCapture "dotnet" @("build", $CliProject, "--nologo", "--verbosity", "minimal")
if ($build.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stdout.log") -Value $build.StdOut
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stderr.log") -Value $build.StdErr
    throw "CLI build failed."
}

$apply = Invoke-ProcessCapture "dotnet" @(
    "run",
    "--no-build",
    "--project",
    $CliProject,
    "--",
    "apply",
    $inputPath,
    $patchPath,
    "-o",
    $outputPath,
    "--json",
    "--report",
    $applyReportPath)
if ($apply.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "apply.stdout.log") -Value $apply.StdOut
    Set-Content -LiteralPath (Join-Path $artifactDir "apply.stderr.log") -Value $apply.StdErr
    throw "Case '$caseId' apply failed."
}

$applyJson = $apply.StdOut | ConvertFrom-Json
$expectedDiagnosticCount = if ($null -ne $manifest.expect.diagnosticCount) { [int] $manifest.expect.diagnosticCount } else { 0 }
$actualDiagnosticCount = @($applyJson.Diagnostics).Count
if ($actualDiagnosticCount -ne $expectedDiagnosticCount) {
    throw "Case '$caseId' diagnostic count mismatch. Expected $expectedDiagnosticCount, found $actualDiagnosticCount."
}

$read = Invoke-ProcessCapture "dotnet" @(
    "run",
    "--no-build",
    "--project",
    $CliProject,
    "--",
    "read",
    $outputPath,
    "--json")
if ($read.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "read.stdout.log") -Value $read.StdOut
    Set-Content -LiteralPath (Join-Path $artifactDir "read.stderr.log") -Value $read.StdErr
    throw "Case '$caseId' readback failed."
}

$readJson = $read.StdOut | ConvertFrom-Json
if ($null -ne $manifest.expect.paragraphs) {
    $expectedParagraphs = @($manifest.expect.paragraphs | ForEach-Object { [string] $_ })
    $actualParagraphs = @($readJson.Paragraphs | ForEach-Object { [string] $_.Text })
    Assert-StringArrayEquals $expectedParagraphs $actualParagraphs "Paragraphs"
}

$summary = [pscustomobject]@{
    CaseId = $caseId
    Success = $true
    ApplyDiagnostics = $actualDiagnosticCount
    OutputParagraphs = @($readJson.Paragraphs).Count
}
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath

Write-Host "case '$caseId': OK"
Write-Host "artifact: $summaryPath"
