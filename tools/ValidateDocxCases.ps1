param(
    [switch] $AllowUncovered
)

$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$CaseRoot = Join-Path $RepoRoot "edit-cases/cases"
$FamilyRoot = Join-Path $RepoRoot "edit-cases/families"
$CheckScript = Join-Path $RepoRoot "tools/CheckDocxCase.ps1"

$cases = @()
if (Test-Path -LiteralPath $CaseRoot) {
    $cases = @(Get-ChildItem -LiteralPath $CaseRoot -Filter "*.json" -File | Sort-Object Name)
}

if ($cases.Count -eq 0) {
    if ($AllowUncovered) {
        Write-Host "no public edit cases found"
        exit 0
    }

    throw "No public edit cases found under edit-cases/cases/."
}

$failed = @()
foreach ($case in $cases) {
    Write-Host "validating $($case.Name)"
    & powershell -NoProfile -ExecutionPolicy Bypass -File $CheckScript -Case $case.FullName
    if ($LASTEXITCODE -ne 0) {
        $failed += $case.Name
    }
}

if ($failed.Count -ne 0) {
    throw "Failed public edit cases: $($failed -join ', ')"
}

$familyRuns = 0
if (Test-Path -LiteralPath $FamilyRoot) {
    foreach ($family in @(Get-ChildItem -LiteralPath $FamilyRoot -Filter "*.json" -File | Sort-Object Name)) {
        $manifest = Get-Content -LiteralPath $family.FullName -Raw | ConvertFrom-Json
        $modes = @($manifest.trackChangesModes | ForEach-Object { [string] $_ })
        if ($modes.Count -eq 0) {
            continue
        }

        foreach ($caseId in @($manifest.cases | ForEach-Object { [string] $_ })) {
            foreach ($mode in $modes) {
                Write-Host "validating $caseId [$mode]"
                $caseArgs = @(
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    $CheckScript,
                    "-Case",
                    $caseId,
                    "-TrackChangesOverride",
                    $mode,
                    "-VariantId",
                    $mode)
                if ($mode -eq "off" -or $mode -eq "preserve") {
                    $caseArgs += "-ExpectNoChangeSummary"
                }

                & powershell $caseArgs
                if ($LASTEXITCODE -ne 0) {
                    $failed += "$caseId [$mode]"
                }
                else {
                    $familyRuns++
                }
            }
        }
    }
}

if ($failed.Count -ne 0) {
    throw "Failed public edit case variants: $($failed -join ', ')"
}

Write-Host "validated $($cases.Count) public edit case(s) and $familyRuns variant run(s)"
