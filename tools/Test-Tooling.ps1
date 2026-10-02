# Tracked offline tooling tests for process capture and challenge verdicts.
# Run locally: pwsh -NoProfile -ExecutionPolicy Bypass -File tools/Test-Tooling.ps1
# CI invokes this file once. Containment and summary-sanitization checks stay as
# separate workflow steps because they assert different boundaries.
$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
. (Join-Path $PSScriptRoot "DocxCaseCommon.ps1")
$failures = 0
function Assert-Tooling([string] $Name, [bool] $Condition, [string] $Detail) {
    if (-not $Condition) {
        Write-Host ("TOOLING-FAIL " + $Name + " " + $Detail)
        $script:failures++
    }
}
$floodChild = '1..4000 | ForEach-Object { [Console]::Out.WriteLine($o); [Console]::Error.WriteLine($e) }'
$floodChild = '$o = "x" * 200; $e = "e" * 200; ' + $floodChild
$flood = Invoke-ProcessCapture "pwsh" @("-NoProfile", "-Command", $floodChild)
Assert-Tooling "flood-exit" ($flood.ExitCode -eq 0) ""
Assert-Tooling "flood-out" ($flood.StdOut.Length -eq 808000) ""
Assert-Tooling "flood-err" ($flood.StdErr.Length -eq 808000) ""
$echoScript = Join-Path ([System.IO.Path]::GetTempPath()) ("tooling-echo-" + [Guid]::NewGuid().ToString("N") + ".ps1")
Set-Content -LiteralPath $echoScript -Value '$args | ConvertTo-Json -Compress' -Encoding UTF8
try {
    $want = @("plain", "with space", "quote""inside", "trailing-backslash\", "", "both ""\ mix")
    $echoed = Invoke-ProcessCapture "pwsh" (@("-NoProfile", "-File", $echoScript) + $want)
    Assert-Tooling "echo-exit" ($echoed.ExitCode -eq 0) ""
    $got = @($echoed.StdOut | ConvertFrom-Json)
    Assert-Tooling "echo-count" ($got.Count -eq $want.Count) ""
    for ($i = 0; $i -lt $want.Count; $i++) {
        Assert-Tooling ("echo-arg-" + $i) ($got[$i] -ceq $want[$i]) ""
    }
    $stdinEcho = Invoke-ProcessCapture "pwsh" @("-NoProfile", "-Command", "[Console]::In.ReadToEnd()") -StandardInput "first"
    Assert-Tooling "stdin-exit" ($stdinEcho.ExitCode -eq 0) ""
    Assert-Tooling "stdin-content" ($stdinEcho.StdOut.Trim() -ceq "first") ""
    $bigPayload = "x" * 1000000
    $slowReader = "import sys,time;time.sleep(3);sys.stdout.write(sys.stdin.read())"
    $slowWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $slowStdin = Invoke-ProcessCapture "python" @("-c", $slowReader) -StandardInput $bigPayload -TimeoutSeconds 1
    $slowWatch.Stop()
    Assert-Tooling "slow-stdin-timeout" ($slowStdin.TimedOut) ""
    Assert-Tooling "slow-stdin-exit" ($slowStdin.ExitCode -eq -1) ""
    Assert-Tooling "slow-stdin-deadline" ($slowWatch.Elapsed.TotalSeconds -lt 15) ""
    $sleepWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $sleepy = Invoke-ProcessCapture "pwsh" @("-NoProfile", "-Command", "Start-Sleep 60") 3
    $sleepWatch.Stop()
    Assert-Tooling "timeout-flag" ($sleepy.TimedOut) ""
    Assert-Tooling "timeout-exit" ($sleepy.ExitCode -eq -1) ""
    Assert-Tooling "timeout-cleanup" (-not $sleepy.CleanupFailed) ""
    $pipeChild = "Start-Process pwsh -ArgumentList ""-NoProfile"",""-Command"",""Start-Sleep 30""; Start-Sleep 60"
    $pipeWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $pipeHeld = Invoke-ProcessCapture "pwsh" @("-NoProfile", "-Command", $pipeChild) 2
    $pipeWatch.Stop()
    Assert-Tooling "pipe-timeout" ($pipeHeld.TimedOut) ""
    Assert-Tooling "pipe-exit" ($pipeHeld.ExitCode -eq -1) ""
    Assert-Tooling "pipe-deadline" ($pipeWatch.Elapsed.TotalSeconds -lt 20) ""
    Assert-Tooling "pipe-cleanup" (-not $pipeHeld.CleanupFailed) ""
}
finally {
    Remove-Item -LiteralPath $echoScript -Force -ErrorAction SilentlyContinue
}
$finOk = [pscustomobject]@{ completed = $true; output_docx = "a.docx" }
$finNo = [pscustomobject]@{ completed = $false; output_docx = $null }
$finOut = [pscustomobject]@{ completed = $true; output_docx = "b.docx" }
$okPost = @([pscustomobject]@{ ReadExitCode = 0; ChangesExitCode = 0 })
$badPost = @([pscustomobject]@{ ReadExitCode = 1; ChangesExitCode = 0 })
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true -SemanticEvidence "verified"
Assert-Tooling "pass-verdict" ($r.Verdict -eq "passed") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "unverified-verdict" ($r.Verdict -eq "unverified") ""
Assert-Tooling "unverified-reason" (($r.FailureReasons -join ",") -eq "unverified-semantics") ""
$finStop = [pscustomobject]@{ completed = $false; output_docx = "a.docx" }
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finStop -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "incomplete-verdict" ($r.Verdict -eq "incomplete") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @() -PostChecks @() -CommandCount 5 -UsedDocxEdit $true -SemanticEvidence "verified"
Assert-Tooling "nooutput-reason" (($r.FailureReasons -join ",") -eq "no-output") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOut -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "mismatch-reason" (($r.FailureReasons -join ",") -eq "output-mismatch") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $badPost -CommandCount 5 -UsedDocxEdit $true -SemanticEvidence "verified"
Assert-Tooling "postcheck-reason" (($r.FailureReasons -join ",") -eq "postcheck-failed") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $false -OutputDocx @() -PostChecks @() -InputModified $true -ForbiddenInspection $true -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "multi-reason" (($r.FailureReasons -join ",") -eq "forbidden-inspection,input-mutated") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finNo -RequiresOutput $true -Applicability "needs repeated placements" -OutputDocx @() -PostChecks @() -CommandCount 3 -UsedDocxEdit $true -FeatureEvidence "absent"
Assert-Tooling "na-verdict" ($r.Verdict -eq "not-applicable") ""
Assert-Tooling "na-reason" (($r.FailureReasons -join ",") -eq "feature-absent") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finNo -RequiresOutput $true -Applicability "needs repeated placements" -OutputDocx @() -PostChecks @() -CommandCount 3 -UsedDocxEdit $true -FeatureEvidence "present"
Assert-Tooling "feature-present-blocks-na" ($r.Verdict -eq "failed") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 0 -UsedDocxEdit $false -SemanticEvidence "verified"
Assert-Tooling "no-command-evidence" (($r.FailureReasons -join ",") -eq "no-task-evidence") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true -SemanticEvidence "mismatch"
Assert-Tooling "semantic-mismatch" (($r.FailureReasons -join ",") -eq "semantic-mismatch") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $null -FinalParseError "final-missing" -RequiresOutput $false -OutputDocx @() -PostChecks @() -CommandCount 3 -UsedDocxEdit $true
Assert-Tooling "unevaluated-verdict" ($r.Verdict -eq "unevaluated") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $null -FinalParseError "final-missing" -RequiresOutput $false -OutputDocx @() -PostChecks @() -InputModified $true -CommandCount 0
Assert-Tooling "missing-final-hides-nothing" ($r.Verdict -eq "failed") ""
$r = Get-ChallengeOutcome -ProcessOk $false -Final $null -RequiresOutput $true -OutputDocx @() -PostChecks @() -CommandCount 0
Assert-Tooling "procfail-verdict" ($r.Verdict -eq "failed") ""
$r = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true -EventParseErrors 2 -SemanticEvidence "verified"
Assert-Tooling "event-errors" (($r.FailureReasons -join ",") -eq "event-parse-errors") ""
$undeclared = Get-ChallengeOutcome -ProcessOk $true -Final ([pscustomobject]@{ completed = $true }) -RequiresOutput $true -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true -SemanticEvidence "verified"
Assert-Tooling "undeclared-output" (($undeclared.FailureReasons -join ",") -eq "undeclared-output") ""
$sharedRead = [pscustomobject]@{ Images = @([pscustomobject]@{ PartName = "a" }, [pscustomobject]@{ PartName = "a" }, [pscustomobject]@{ PartName = "b" }) }
$soloRead = [pscustomobject]@{ Images = @([pscustomobject]@{ PartName = "a" }, [pscustomobject]@{ PartName = "b" }) }
Assert-Tooling "shared-media-present" ((Test-SharedMediaFeature $sharedRead) -eq "present") ""
Assert-Tooling "shared-media-absent" ((Test-SharedMediaFeature $soloRead) -eq "absent") ""
Assert-Tooling "shared-media-unknown" ((Test-SharedMediaFeature $null) -eq "") ""
$spanRead = [pscustomobject]@{ Bookmarks = @([pscustomobject]@{ StartTargetId = "M.P0001"; EndTargetId = "M.P0003" }); Fields = @() }
$flatRead = [pscustomobject]@{ Bookmarks = @([pscustomobject]@{ StartTargetId = "M.P0001"; EndTargetId = "M.P0001" }); Fields = @() }
$fieldRead = [pscustomobject]@{ Bookmarks = @(); Fields = @([pscustomobject]@{ Kind = "complex" }) }
$bareRead = [pscustomobject]@{ Bookmarks = @(); Fields = @() }
Assert-Tooling "range-span-present" ((Test-MultiparagraphRangeFeature $spanRead) -eq "present") ""
Assert-Tooling "range-flat-absent" ((Test-MultiparagraphRangeFeature $flatRead) -eq "absent") ""
Assert-Tooling "range-field-present" ((Test-MultiparagraphRangeFeature $fieldRead) -eq "present") ""
Assert-Tooling "range-bare-absent" ((Test-MultiparagraphRangeFeature $bareRead) -eq "absent") ""
$imgBefore = [pscustomobject]@{ Images = @([pscustomobject]@{ Id = "M.I0001"; Description = "A" }, [pscustomobject]@{ Id = "M.I0002"; Description = "B" }) }
$imgAfter = [pscustomobject]@{ Images = @([pscustomobject]@{ Id = "M.I0001"; Description = "C" }, [pscustomobject]@{ Id = "M.I0002"; Description = "B" }) }
Assert-Tooling "single-image-verified" ((Test-SingleImageChange $imgBefore $imgAfter) -eq "verified") ""
Assert-Tooling "single-image-same" ((Test-SingleImageChange $imgBefore $imgBefore) -eq "mismatch") ""
$paraBefore = [pscustomobject]@{ Paragraphs = @([pscustomobject]@{ Text = "One" }); Bookmarks = @([pscustomobject]@{ StartTargetId = "M.P0001"; EndTargetId = "M.P0002" }); Fields = @() }
$paraAfter = [pscustomobject]@{ Paragraphs = @([pscustomobject]@{ Text = "Two" }); Bookmarks = @([pscustomobject]@{ StartTargetId = "M.P0001"; EndTargetId = "M.P0002" }); Fields = @() }
Assert-Tooling "markup-preserved-verified" ((Test-MarkupPreserved $paraBefore $paraAfter @("bookmark")) -eq "verified") ""
Assert-Tooling "markup-preserved-same" ((Test-MarkupPreserved $paraBefore $paraBefore @("bookmark")) -eq "mismatch") ""
Assert-Tooling "output-differs-verified" ((Test-OutputDiffers $paraBefore $paraAfter) -eq "verified") ""
Assert-Tooling "output-differs-same" ((Test-OutputDiffers $paraBefore $paraBefore) -eq "mismatch") ""
$missingPath = Join-Path ([System.IO.Path]::GetTempPath()) ("tooling-final-" + [Guid]::NewGuid().ToString("N") + ".json")
$missing = Read-FinalResponse $missingPath $true
Assert-Tooling "final-missing" ($missing.ParseError -eq "final-missing") ""
$badPath = Join-Path ([System.IO.Path]::GetTempPath()) ("tooling-final-" + [Guid]::NewGuid().ToString("N") + ".json")
Set-Content -LiteralPath $badPath -Value "{not json" -Encoding UTF8
$bad = Read-FinalResponse $badPath $true
Assert-Tooling "final-unparseable" ($bad.ParseError -eq "final-unparseable") ""
Set-Content -LiteralPath $badPath -Value "{completed: ""yes""}" -Encoding UTF8
$invalid = Read-FinalResponse $badPath $true
Assert-Tooling "final-invalid" ($invalid.ParseError -eq "final-invalid") ""
Remove-Item -LiteralPath $badPath -Force -ErrorAction SilentlyContinue
if ($failures -eq 0) { Write-Host "TOOLING-TESTS-OK" } else { throw ($failures.ToString() + " tooling failures") }
