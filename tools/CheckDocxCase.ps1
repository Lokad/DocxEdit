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

function Resolve-PublicFixturePath([string] $FixtureValue) {
    if ([string]::IsNullOrWhiteSpace($FixtureValue)) {
        throw "Fixture path must not be empty."
    }

    $candidate = if ([System.IO.Path]::IsPathRooted($FixtureValue)) {
        $FixtureValue
    } else {
        Join-Path $RepoRoot $FixtureValue
    }

    $resolved = (Resolve-Path -LiteralPath $candidate).Path
    $fixtureRoot = [System.IO.Path]::GetFullPath((Join-Path $RepoRoot "edit-cases/fixtures")).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $resolvedFull = [System.IO.Path]::GetFullPath($resolved)
    if (-not $resolvedFull.StartsWith($fixtureRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Public fixture inputs must live under edit-cases/fixtures/."
    }

    return $resolved
}

function Get-ManifestText($Value) {
    if ($Value -is [System.Array]) {
        return ($Value -join [Environment]::NewLine)
    }

    return [string] $Value
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

function New-DocxFromBody([string] $Path, [string] $BodyXml, [string] $HeaderXml, [string] $FooterXml) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path
    }

    $hasHeader = -not [string]::IsNullOrWhiteSpace($HeaderXml)
    $hasFooter = -not [string]::IsNullOrWhiteSpace($FooterXml)
    $headerOverride = if ($hasHeader) { '  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' } else { '' }
    $footerOverride = if ($hasFooter) { '  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' } else { '' }
    $relationships = @('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    if ($hasHeader) {
        $relationships += '  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    }

    if ($hasFooter) {
        $relationships += '  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    }

    $relationships += '</Relationships>'
    $documentRelationships = $relationships -join [Environment]::NewLine

    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipEntry $archive "[Content_Types].xml" @"
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
$headerOverride
$footerOverride
</Types>
"@
        Add-ZipEntry $archive "_rels/.rels" @"
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@
        Add-ZipEntry $archive "word/_rels/document.xml.rels" $documentRelationships
        Add-ZipEntry $archive "word/document.xml" @"
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
$BodyXml
  </w:body>
</w:document>
"@
        if ($hasHeader) {
            Add-ZipEntry $archive "word/header1.xml" @"
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
$HeaderXml
</w:hdr>
"@
        }

        if ($hasFooter) {
            Add-ZipEntry $archive "word/footer1.xml" @"
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
$FooterXml
</w:ftr>
"@
        }
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

$fixtureValue = Get-ObjectProperty $manifest.input "fixture"
$bodyXml = Get-ManifestText $manifest.input.bodyXml
$headerXml = Get-ManifestText (Get-ObjectProperty $manifest.input "headerXml")
$footerXml = Get-ManifestText (Get-ObjectProperty $manifest.input "footerXml")
$patchText = Get-ManifestText $manifest.patch
if ([string]::IsNullOrWhiteSpace($patchText)) {
    throw "Case '$caseId' must define patch."
}

if ($null -eq $fixtureValue -and [string]::IsNullOrWhiteSpace($bodyXml)) {
    throw "Case '$caseId' must define input.bodyXml or input.fixture."
}

$runId = [DateTimeOffset]::UtcNow.ToString("yyyyMMddTHHmmssZ")
$artifactDir = Join-Path $ArtifactRoot (Join-Path $caseId $runId)
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null

$inputPath = Join-Path $artifactDir "input.docx"
$patchPath = Join-Path $artifactDir "edit.docxpatch"
$outputPath = Join-Path $artifactDir "output.docx"
$applyReportPath = Join-Path $artifactDir "apply-report.json"
$summaryPath = Join-Path $artifactDir "summary.json"

$assets = Get-ObjectProperty $manifest "assets"
if ($null -ne $assets) {
    $assetDir = Join-Path $artifactDir "assets"
    New-Item -ItemType Directory -Force -Path $assetDir | Out-Null
    foreach ($property in @($assets.PSObject.Properties)) {
        $assetPath = Join-Path $assetDir ([string] $property.Name)
        [System.IO.File]::WriteAllText($assetPath, [string] $property.Value, [System.Text.Encoding]::UTF8)
        $patchText = $patchText.Replace("{{asset:$($property.Name)}}", $assetPath)
    }
}

if ($null -ne $fixtureValue) {
    Copy-Item -LiteralPath (Resolve-PublicFixturePath ([string] $fixtureValue)) -Destination $inputPath
}
else {
    New-DocxFromBody $inputPath $bodyXml $headerXml $footerXml
}
Set-Content -LiteralPath $patchPath -Value $patchText

$build = Invoke-ProcessCapture "dotnet" @("build", $CliProject, "--nologo", "--verbosity", "minimal")
if ($build.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stdout.log") -Value $build.StdOut
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stderr.log") -Value $build.StdErr
    throw "CLI build failed."
}

$applyArgs = @(
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
$applyOptions = Get-ObjectProperty $manifest "applyOptions"
if ($null -ne $applyOptions) {
    $trackChanges = Get-ObjectProperty $applyOptions "trackChanges"
    if ($null -ne $trackChanges) {
        $applyArgs += @("--track-changes", [string] $trackChanges)
    }

    $author = Get-ObjectProperty $applyOptions "author"
    if ($null -ne $author) {
        $applyArgs += @("--author", [string] $author)
    }

    $timestampUtc = Get-ObjectProperty $applyOptions "timestampUtc"
    if ($null -ne $timestampUtc) {
        $applyArgs += @("--timestamp-utc", [string] $timestampUtc)
    }
}

$apply = Invoke-ProcessCapture "dotnet" $applyArgs
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

if ($null -ne $manifest.expect.tableCells) {
    $expectedCells = @($manifest.expect.tableCells | ForEach-Object { [string] $_ })
    $actualCells = @($readJson.Tables | ForEach-Object { $_.Cells } | ForEach-Object { [string] $_.Text })
    Assert-StringArrayEquals $expectedCells $actualCells "Table cells"
}

if ($null -ne $manifest.expect.sections) {
    $expectedSections = @($manifest.expect.sections | ForEach-Object { [string] $_ })
    $actualSections = @($readJson.Sections | ForEach-Object { "$($_.Columns)|$($_.Orientation)" })
    Assert-StringArrayEquals $expectedSections $actualSections "Sections"
}

if ($null -ne $manifest.expect.images) {
    $expectedImages = @($manifest.expect.images | ForEach-Object { [string] $_ })
    $actualImages = @($readJson.Images | ForEach-Object { [string] $_.PartName })
    Assert-StringArrayEquals $expectedImages $actualImages "Images"
}

if ($null -ne $manifest.expect.allStoryParagraphs) {
    $allStoryRead = Invoke-ProcessCapture "dotnet" @(
        "run",
        "--no-build",
        "--project",
        $CliProject,
        "--",
        "read",
        $outputPath,
        "--all-stories",
        "--json")
    if ($allStoryRead.ExitCode -ne 0) {
        Set-Content -LiteralPath (Join-Path $artifactDir "read-all-stories.stdout.log") -Value $allStoryRead.StdOut
        Set-Content -LiteralPath (Join-Path $artifactDir "read-all-stories.stderr.log") -Value $allStoryRead.StdErr
        throw "Case '$caseId' all-story readback failed."
    }

    $allStoryReadJson = $allStoryRead.StdOut | ConvertFrom-Json
    $expectedAllStoryParagraphs = @($manifest.expect.allStoryParagraphs | ForEach-Object { [string] $_ })
    $actualAllStoryParagraphs = @($allStoryReadJson.Paragraphs | ForEach-Object { [string] $_.Text })
    Assert-StringArrayEquals $expectedAllStoryParagraphs $actualAllStoryParagraphs "All-story paragraphs"
}

if ($null -ne $manifest.expect.changeSummary) {
    $changes = Invoke-ProcessCapture "dotnet" @(
        "run",
        "--no-build",
        "--project",
        $CliProject,
        "--",
        "changes",
        $outputPath,
        "--json")
    if ($changes.ExitCode -ne 0) {
        Set-Content -LiteralPath (Join-Path $artifactDir "changes.stdout.log") -Value $changes.StdOut
        Set-Content -LiteralPath (Join-Path $artifactDir "changes.stderr.log") -Value $changes.StdErr
        throw "Case '$caseId' changes readback failed."
    }

    $changesJson = $changes.StdOut | ConvertFrom-Json
    $actualSummary = @{}
    foreach ($summaryItem in @($changesJson.Summary)) {
        $actualSummary[[string] $summaryItem.Type] = [int] $summaryItem.Count
    }

    foreach ($property in @($manifest.expect.changeSummary.PSObject.Properties)) {
        $key = [string] $property.Name
        $expected = [int] $property.Value
        $actual = if ($actualSummary.ContainsKey($key)) { $actualSummary[$key] } else { 0 }
        if ($actual -ne $expected) {
            throw "Change summary mismatch for '$key'. Expected $expected, found $actual."
        }
    }
}

$summary = [pscustomobject]@{
    CaseId = $caseId
    Success = $true
    ApplyDiagnostics = $actualDiagnosticCount
    OutputParagraphs = @($readJson.Paragraphs).Count
    OutputTables = @($readJson.Tables).Count
    OutputImages = @($readJson.Images).Count
    OutputSections = @($readJson.Sections).Count
}
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath

Write-Host "case '$caseId': OK"
Write-Host "artifact: $summaryPath"
