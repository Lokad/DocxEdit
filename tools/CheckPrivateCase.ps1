param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [switch] $ValidateOnly
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "DocxCaseCommon.ps1")

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PrivateRoot = Join-Path $RepoRoot "private-cases"
$ArtifactRoot = Join-Path $RepoRoot "artifacts/private-edit"
$CliProject = Join-Path $RepoRoot "src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj"

function Resolve-ExistingPath([string] $Path) {
    return (Resolve-Path -LiteralPath $Path).Path
}

function Read-ZipEntryText([System.IO.Compression.ZipArchive] $Archive, [string] $EntryName) {
    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        return $null
    }

    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Add-WmlNamespaces([xml] $Document) {
    $manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
    [void] $manager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    [void] $manager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    [void] $manager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    return ,$manager
}

function Count-XmlNodes([xml] $Document, [System.Xml.XmlNamespaceManager] $Namespaces, [string] $XPath) {
    $nodes = $Document.SelectNodes($XPath, $Namespaces)
    if ($null -eq $nodes) {
        return 0
    }

    return $nodes.Count
}

function Resolve-ZipRelationshipTarget([string] $SourceEntryName, [string] $Target) {
    if ($Target.StartsWith("/", [System.StringComparison]::Ordinal)) {
        return $Target.TrimStart("/")
    }

    $sourceDirectory = [System.IO.Path]::GetDirectoryName($SourceEntryName).Replace('\', '/')
    if ([string]::IsNullOrEmpty($sourceDirectory)) {
        return $Target
    }

    return ($sourceDirectory + "/" + $Target).Replace('\', '/')
}

function Get-RelatedStoryEntries([System.IO.Compression.ZipArchive] $Archive) {
    $relationshipsText = Read-ZipEntryText $Archive "word/_rels/document.xml.rels"
    if ($null -eq $relationshipsText) {
        return @()
    }

    [xml] $relationshipsXml = $relationshipsText
    $manager = [System.Xml.XmlNamespaceManager]::new($relationshipsXml.NameTable)
    [void] $manager.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
    $nodes = $relationshipsXml.SelectNodes("/rel:Relationships/rel:Relationship", $manager)
    $entries = @()
    foreach ($node in @($nodes)) {
        $type = [string] $node.Type
        $targetMode = [string] $node.TargetMode
        if ($targetMode.Equals("External", [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        if ($type -ne "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" -and
            $type -ne "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer") {
            continue
        }

        $target = [string] $node.Target
        if ([string]::IsNullOrWhiteSpace($target)) {
            continue
        }

        $entryName = Resolve-ZipRelationshipTarget "word/document.xml" $target
        if ($null -ne $Archive.GetEntry($entryName)) {
            $entries += $entryName
        }
    }

    return $entries
}

function Get-StorySummary([System.IO.Compression.ZipArchive] $Archive, [string] $EntryName, [bool] $IncludeSections) {
    $storyText = Read-ZipEntryText $Archive $EntryName
    if ($null -eq $storyText) {
        throw "Private input is missing an expected story part."
    }

    [xml] $storyXml = $storyText
    $namespaces = Add-WmlNamespaces $storyXml
    $rootName = $storyXml.DocumentElement.LocalName

    if ($rootName -eq "document") {
        $paragraphXPath = "/w:document/w:body/w:p"
        $tableXPath = "/w:document/w:body/w:tbl"
        $sectionCount = if ($IncludeSections) {
            (Count-XmlNodes $storyXml $namespaces "/w:document/w:body/w:sectPr") +
                (Count-XmlNodes $storyXml $namespaces "/w:document/w:body/w:p/w:pPr/w:sectPr")
        } else {
            0
        }
    }
    elseif ($rootName -eq "hdr") {
        $paragraphXPath = "/w:hdr/w:p"
        $tableXPath = "/w:hdr/w:tbl"
        $sectionCount = 0
    }
    elseif ($rootName -eq "ftr") {
        $paragraphXPath = "/w:ftr/w:p"
        $tableXPath = "/w:ftr/w:tbl"
        $sectionCount = 0
    }
    else {
        $paragraphXPath = "/*/w:p"
        $tableXPath = "/*/w:tbl"
        $sectionCount = 0
    }

    [pscustomobject]@{
        Paragraphs = Count-XmlNodes $storyXml $namespaces $paragraphXPath
        Tables = Count-XmlNodes $storyXml $namespaces $tableXPath
        ScannerVisibleImages = Count-XmlNodes $storyXml $namespaces "//w:drawing//a:blip[@r:embed]"
        Sections = $sectionCount
    }
}

function Get-StructuralSummary([string] $InputPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = [System.IO.Compression.ZipFile]::OpenRead($InputPath)
    try {
        $partCount = @($archive.Entries | Where-Object { -not $_.FullName.EndsWith("/", [System.StringComparison]::Ordinal) }).Count
        $stylesText = Read-ZipEntryText $archive "word/styles.xml"
        $styleCount = 0
        if ($null -ne $stylesText) {
            [xml] $stylesXml = $stylesText
            $styleNamespaces = Add-WmlNamespaces $stylesXml
            $styleCount = Count-XmlNodes $stylesXml $styleNamespaces "//w:style[@w:styleId and (@w:type='paragraph' or @w:type='character' or @w:type='table')]"
        }

        $paragraphCount = 0
        $tableCount = 0
        $imageCount = 0
        $sectionCount = 0
        $storyEntries = @("word/document.xml") + @(Get-RelatedStoryEntries $archive)
        foreach ($storyEntry in @($storyEntries | Select-Object -Unique)) {
            $story = Get-StorySummary $archive $storyEntry ($storyEntry -eq "word/document.xml")
            $paragraphCount += $story.Paragraphs
            $tableCount += $story.Tables
            $imageCount += $story.ScannerVisibleImages
            $sectionCount += $story.Sections
        }

        [pscustomobject]@{
            PackageParts = $partCount
            Paragraphs = $paragraphCount
            Tables = $tableCount
            ScannerVisibleImages = $imageCount
            Sections = $sectionCount
            Styles = $styleCount
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Compare-ExpectedAggregate(
    [System.Collections.Generic.List[string]] $Failures,
    [string] $Name,
    [object] $Actual,
    [object] $Expected) {
    if ($null -eq $Expected) {
        return
    }

    $actualValue = [int64] $Actual
    $expectedValue = [int64] $Expected
    if ($actualValue -ne $expectedValue) {
        [void] $Failures.Add("$Name expected $expectedValue, found $actualValue")
    }
}

function Compare-ExpectedAggregates([object] $Expected, [object] $Structure, [hashtable] $ChangeSummary, [int] $ChangeTotal) {
    $failures = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Expected) {
        return ,$failures
    }

    $expectedStructure = Get-ObjectProperty $Expected "Structure"
    if ($null -ne $expectedStructure) {
        Compare-ExpectedAggregate $failures "structure.PackageParts" $Structure.PackageParts (Get-ObjectProperty $expectedStructure "PackageParts")
        Compare-ExpectedAggregate $failures "structure.Paragraphs" $Structure.Paragraphs (Get-ObjectProperty $expectedStructure "Paragraphs")
        Compare-ExpectedAggregate $failures "structure.Tables" $Structure.Tables (Get-ObjectProperty $expectedStructure "Tables")
        Compare-ExpectedAggregate $failures "structure.ScannerVisibleImages" $Structure.ScannerVisibleImages (Get-ObjectProperty $expectedStructure "ScannerVisibleImages")
        Compare-ExpectedAggregate $failures "structure.Sections" $Structure.Sections (Get-ObjectProperty $expectedStructure "Sections")
        Compare-ExpectedAggregate $failures "structure.Styles" $Structure.Styles (Get-ObjectProperty $expectedStructure "Styles")
    }

    $expectedChanges = Get-ObjectProperty $Expected "Changes"
    if ($null -ne $expectedChanges) {
        Compare-ExpectedAggregate $failures "changes.Total" $ChangeTotal (Get-ObjectProperty $expectedChanges "Total")
        $expectedByType = Get-ObjectProperty $expectedChanges "ByType"
        if ($null -ne $expectedByType) {
            foreach ($property in @($expectedByType.PSObject.Properties)) {
                $key = [string] $property.Name
                $actual = if ($ChangeSummary.ContainsKey($key)) { $ChangeSummary[$key] } else { 0 }
                Compare-ExpectedAggregate $failures "changes.ByType.$key" $actual $property.Value
            }
        }
    }

    return ,$failures
}

function Resolve-PrivateCase([string] $CaseValue) {
    if ([System.IO.Path]::IsPathRooted($CaseValue)) {
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

    $casePath = Resolve-ExistingPath $candidate
    Assert-PrivatePath $casePath "Private case"

    if ([System.IO.Path]::GetExtension($casePath).Equals(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
        $manifest = Get-Content -LiteralPath $casePath -Raw | ConvertFrom-Json
        $caseId = [string] $manifest.id
        if ([string]::IsNullOrWhiteSpace($caseId)) {
            throw "Private manifest must define an id."
        }

        $inputValue = [string] $manifest.input
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            throw "Private manifest must define an input."
        }

        if ([System.IO.Path]::IsPathRooted($inputValue)) {
            $inputPath = Resolve-ExistingPath $inputValue
        }
        else {
            $inputPath = Resolve-ExistingPath (Join-Path (Split-Path -Parent $casePath) $inputValue)
        }

        Assert-PrivatePath $inputPath "Private input"
        $expected = Get-ObjectProperty $manifest "expected"
    }
    else {
        $inputPath = $casePath
        $caseId = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
        $manifestPath = $null
        $expected = $null
        $siblingManifest = [System.IO.Path]::ChangeExtension($inputPath, ".json")
        if (Test-Path -LiteralPath $siblingManifest) {
            $manifestPath = Resolve-ExistingPath $siblingManifest
            Assert-PrivatePath $manifestPath "Private manifest"
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            $expected = Get-ObjectProperty $manifest "expected"
        }
    }

    if ($caseId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw "Private case id must be a single filename-safe path segment."
    }

    if (-not [System.IO.Path]::GetExtension($inputPath).Equals(".docx", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Private input must be a .docx file."
    }

    [pscustomobject]@{
        CaseId = $caseId
        InputPath = $inputPath
        ManifestPath = if ($inputPath -eq $casePath) { $null } else { $casePath }
        Expected = $expected
    }
}

$privateCase = Resolve-PrivateCase $Case
$runId = [DateTimeOffset]::UtcNow.ToString("yyyyMMddTHHmmssZ")
$artifactDir = Join-Path $ArtifactRoot (Join-Path $privateCase.CaseId $runId)
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null

$build = Invoke-ProcessCapture "dotnet" @("build", $CliProject, "--nologo", "--verbosity", "minimal")
if ($build.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stdout.log") -Value $build.StdOut
    Set-Content -LiteralPath (Join-Path $artifactDir "build.stderr.log") -Value $build.StdErr
    throw "CLI build failed. See private artifact logs."
}

$structure = Get-StructuralSummary $privateCase.InputPath

$changes = Invoke-ProcessCapture "dotnet" @(
    "run",
    "--no-build",
    "--project",
    $CliProject,
    "--",
    "changes",
    $privateCase.InputPath,
    "--json")
if ($changes.ExitCode -ne 0) {
    Set-Content -LiteralPath (Join-Path $artifactDir "changes.stderr.log") -Value $changes.StdErr
    throw "Tracked-change validation failed. See private artifact logs."
}

$changesJson = $changes.StdOut | ConvertFrom-Json
$diagnostics = @($changesJson.Diagnostics)
$changeRecords = @($changesJson.Changes)
$changeSummary = @{}
foreach ($summary in @($changesJson.Summary)) {
    $changeSummary[[string] $summary.Type] = [int] $summary.Count
}

$aggregateFailures = Compare-ExpectedAggregates $privateCase.Expected $structure $changeSummary $changeRecords.Count
$success = [bool] $changesJson.Success -and $diagnostics.Count -eq 0 -and $aggregateFailures.Count -eq 0
$summaryObject = [pscustomobject]@{
    CaseId = $privateCase.CaseId
    RunId = $runId
    ValidateOnly = [bool] $ValidateOnly
    Success = $success
    InputBytes = (Get-Item -LiteralPath $privateCase.InputPath).Length
    Structure = $structure
    Changes = [pscustomobject]@{
        Total = $changeRecords.Count
        ByType = $changeSummary
    }
    Diagnostics = @($diagnostics | ForEach-Object {
        [pscustomobject]@{
            Code = $_.Code
            Severity = $_.Severity
            Message = $_.Message
        }
    })
    AggregateFailures = @($aggregateFailures)
}

$summaryPath = Join-Path $artifactDir "summary.json"
$summaryObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath

Write-Host "private case '$($privateCase.CaseId)': $(if ($success) { 'OK' } else { 'FAILED' })"
Write-Host "artifact: $(Get-RepoRelativePath $summaryPath)"
Write-Host "structure: parts=$($structure.PackageParts) paragraphs=$($structure.Paragraphs) tables=$($structure.Tables) images=$($structure.ScannerVisibleImages) sections=$($structure.Sections) styles=$($structure.Styles)"
Write-Host "changes: total=$($changeRecords.Count)"
foreach ($key in @($changeSummary.Keys | Sort-Object)) {
    Write-Host "  $key=$($changeSummary[$key])"
}
Write-Host "diagnostics: count=$($diagnostics.Count)"
foreach ($failure in @($aggregateFailures)) {
    Write-Host "aggregate-check: $failure"
}

if (-not $success) {
    exit 1
}
