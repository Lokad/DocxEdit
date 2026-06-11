param(
    [switch] $AllowUncovered
)

$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$CaseRoot = Join-Path $RepoRoot "edit-cases/cases"
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

Write-Host "validated $($cases.Count) public edit case(s)"
