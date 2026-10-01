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
    Assert-Tooling "slow-stdin-deadline" ($slowWatch.Elapsed.TotalSeconds -lt 2.5) ""
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
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -Applicability "" -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "pass-verdict" ($outcome.Verdict -eq "passed") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finNo -RequiresOutput $true -Applicability "" -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "incomplete-verdict" ($outcome.Verdict -eq "failed") ""
Assert-Tooling "incomplete-reason" (($outcome.FailureReasons -join ",") -eq "final-incomplete") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -Applicability "" -OutputDocx @() -PostChecks @() -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "no-output-reason" (($outcome.FailureReasons -join ",") -eq "no-output") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finOut -RequiresOutput $true -Applicability "" -OutputDocx @("a.docx") -PostChecks $okPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "mismatch-reason" (($outcome.FailureReasons -join ",") -eq "output-mismatch") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $true -Applicability "" -OutputDocx @("a.docx") -PostChecks $badPost -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "postcheck-reason" (($outcome.FailureReasons -join ",") -eq "postcheck-failed") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finOk -RequiresOutput $false -Applicability "" -OutputDocx @() -PostChecks @() -InputModified $true -ForbiddenInspection $true -CommandCount 5 -UsedDocxEdit $true
Assert-Tooling "multi-reason" (($outcome.FailureReasons -join ",") -eq "forbidden-inspection,input-mutated") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $finNo -RequiresOutput $true -Applicability "needs repeated placements" -OutputDocx @() -PostChecks @() -CommandCount 3 -UsedDocxEdit $true
Assert-Tooling "na-verdict" ($outcome.Verdict -eq "not-applicable") ""
Assert-Tooling "na-reason" (($outcome.FailureReasons -join ",") -eq "feature-absent") ""
$outcome = Get-ChallengeOutcome -ProcessOk $true -Final $null -FinalParseError "final-missing" -RequiresOutput $true -Applicability "" -OutputDocx @() -PostChecks @() -CommandCount 0
Assert-Tooling "unevaluated-verdict" ($outcome.Verdict -eq "unevaluated") ""
$outcome = Get-ChallengeOutcome -ProcessOk $false -Final $null -RequiresOutput $true -Applicability "" -OutputDocx @() -PostChecks @() -CommandCount 0
Assert-Tooling "procfail-verdict" ($outcome.Verdict -eq "failed") ""
Assert-Tooling "procfail-reason" (($outcome.FailureReasons -join ",") -eq "no-output,process-failed") ""
$missingPath = Join-Path ([System.IO.Path]::GetTempPath()) ("tooling-final-" + [Guid]::NewGuid().ToString("N") + ".json")
$missing = Read-FinalResponse $missingPath $true
Assert-Tooling "final-missing" ($missing.ParseError -eq "final-missing") ""
$badPath = Join-Path ([System.IO.Path]::GetTempPath()) ("tooling-final-" + [Guid]::NewGuid().ToString("N") + ".json")
Set-Content -LiteralPath $badPath -Value "{not json" -Encoding UTF8
$bad = Read-FinalResponse $badPath $true
Assert-Tooling "final-unparseable" ($bad.ParseError -eq "final-unparseable") ""
Set-Content -LiteralPath $badPath -Value "{""completed"": ""yes""}" -Encoding UTF8
$invalid = Read-FinalResponse $badPath $true
Assert-Tooling "final-invalid" ($invalid.ParseError -eq "final-invalid") ""
Remove-Item -LiteralPath $badPath -Force -ErrorAction SilentlyContinue
if ($failures -eq 0) { Write-Host "TOOLING-TESTS-OK" } else { throw ($failures.ToString() + " tooling failures") }
