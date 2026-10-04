param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../artifacts/equation-smoke'),
    [switch]$WordWorker,
    [string]$DocumentPath,
    [string]$PdfPath,
    [string]$ProcessIdPath
)

$ErrorActionPreference = 'Stop'

# Separate STA worker: the parent imposes a deadline and only terminates the
# Word process whose HWND this worker owns. No user documents are opened.
if ($WordWorker) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EquationWordWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
    $wordApp = $null
    $wordDocuments = $null
    $wordDocument = $null
    try {
        Write-Output 'Creating Word application'
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.DisplayAlerts = 0
        $wordApp.AutomationSecurity = 3
        $wordDocuments = $wordApp.Documents
        Write-Output 'Opening synthetic equation document'
        $wordDocument = $wordDocuments.Open($DocumentPath, $false, $true, $false)
        $wordWindow = $wordDocument.ActiveWindow
        [uint32]$wordProcessId = 0
        [void][EquationWordWindow]::GetWindowThreadProcessId([IntPtr]$wordWindow.Hwnd, [ref]$wordProcessId)
        Set-Content -LiteralPath $ProcessIdPath -Value $wordProcessId
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordWindow)
        $maths = $wordDocument.OMaths
        Write-Output ('Native Word equations: ' + $maths.Count)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($maths)
        $wordDocument.ExportAsFixedFormat($PdfPath, 17)
        Write-Output 'PDF exported'
        $roundTripPath = [IO.Path]::ChangeExtension($DocumentPath, 'word.docx')
        $wordDocument.SaveAs2([ref]$roundTripPath)
        Write-Output 'Word copy saved'
        $wordDocument.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDocument)
        $wordDocument = $null
    }
    finally {
        if ($null -ne $wordDocument) { $wordDocument.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDocument) }
        if ($null -ne $wordDocuments) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDocuments) }
        if ($null -ne $wordApp) { $wordApp.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) }
    }
    exit
}

$repoPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$cliProject = Join-Path $repoPath 'src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj'
& dotnet build $cliProject -c Release -v quiet
if ($LASTEXITCODE -ne 0) { throw 'CLI build failed.' }
$cliDll = Join-Path $repoPath 'src/Lokad.DocxEdit.Cli/bin/Release/net10.0/Lokad.DocxEdit.Cli.dll'
function Invoke-Docx([string[]]$CommandArgs) {
    $result = & dotnet $cliDll @CommandArgs
    if ($LASTEXITCODE -ne 0) { throw ('docxedit failed: ' + ($CommandArgs -join ' ')) }
    return $result
}

$blankPath = Join-Path $OutputDirectory 'blank.docx'
$generatedPath = Join-Path $OutputDirectory 'equations-generated.docx'
$editedPath = Join-Path $OutputDirectory 'equations-edited.docx'
$createPatchPath = Join-Path $OutputDirectory 'create.docxpatch'
$editPatchPath = Join-Path $OutputDirectory 'edit.docxpatch'
Invoke-Docx @('create', '--output', $blankPath) | Out-Host
$examples = @(
    @('Quadratic formula (replaced in the edited document)', 'x^2'),
    @('Subscripts, superscripts, and Greek letters', 'x_i^2+\alpha\beta=\gamma'),
    @('Square root and cube root', '\sqrt{1+x^2}+\sqrt[3]{y}'),
    @('Summation with limits', '\sum_{i=1}^{n}{i^2}=\frac{n(n+1)(2n+1)}{6}'),
    @('Integral with limits', '\int_0^1{\frac{1}{1+x^2}}\,\mathrm{d}x=\frac{\pi}{4}'),
    @('Matrix', 'A=\begin{pmatrix}a&b\\c&d\end{pmatrix}'),
    @('Scalable delimiters', '\left(\frac{a+b}{c}\right)^2\geq0'),
    @('Cases and upright text', 'f(x)=\begin{cases}x^2&\text{if }x\geq0\\-x&\text{otherwise}\end{cases}')
)
$patch = [Text.StringBuilder]::new()
[void]$patch.AppendLine("docxpatch 1`nop replace-paragraph`ntarget M.P0001`ntext Native Word equations - smoke test`nstyle Heading1`nend")
for ($i = $examples.Count - 1; $i -ge 0; $i--) {
    [void]$patch.AppendLine("op insert-after`ntarget M.P0001`ntext $($examples[$i][0])`nas label$i`nend")
    [void]$patch.AppendLine("op insert-equation`ntarget @label$i`nlatex $($examples[$i][1])`nend")
}
[void]$patch.AppendLine("op insert-after`ntarget M.P0001`ntext Inline formula: `nas inlineLabel`nend")
[void]$patch.AppendLine('op insert-equation' + "`ntarget @inlineLabel`nplacement inline`nlatex \,E=mc^2`nend")
[void]$patch.AppendLine("op insert-equation`ntarget M.P0001`nlatex TEMPORARY`nend")
[IO.File]::WriteAllText($createPatchPath, $patch.ToString())
Invoke-Docx @('apply', $blankPath, $createPatchPath, '--output', $generatedPath) | Out-Host
$read = (Invoke-Docx @('read', $generatedPath, '--json', '--max-text', '100000')) -join "`n" | ConvertFrom-Json
$quadratic = @($read.Equations | Where-Object Text -eq 'x2')[0]
$temporary = @($read.Equations | Where-Object Text -eq 'TEMPORARY')[0]
$edit = "docxpatch 1`nop replace-equation`ntarget $($quadratic.Id)`nexpect-hash $($quadratic.ContentHash)`nlatex x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`nend`nop delete-equation`ntarget $($temporary.Id)`nexpect-hash $($temporary.ContentHash)`nend"
[IO.File]::WriteAllText($editPatchPath, $edit)
Invoke-Docx @('check', $generatedPath, $editPatchPath) | Out-Host
Invoke-Docx @('apply', $generatedPath, $editPatchPath, '--output', $editedPath) | Out-Host

foreach ($docPath in @($generatedPath, $editedPath)) {
    $pdf = [IO.Path]::ChangeExtension($docPath, 'pdf')
    $log = [IO.Path]::ChangeExtension($docPath, 'word.log')
    $errorLog = [IO.Path]::ChangeExtension($docPath, 'word-error.log')
    $pidPath = [IO.Path]::ChangeExtension($docPath, 'word.pid')
    if (Test-Path -LiteralPath $pidPath) { Remove-Item -LiteralPath $pidPath }
    $workerArguments = @('-NoProfile', '-STA', '-File', ('"' + $PSCommandPath + '"'), '-WordWorker',
        '-DocumentPath', ('"' + $docPath + '"'), '-PdfPath', ('"' + $pdf + '"'), '-ProcessIdPath', ('"' + $pidPath + '"'))
    $worker = Start-Process powershell.exe -ArgumentList $workerArguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $log -RedirectStandardError $errorLog
    $workerStartTime = $worker.StartTime
    if (!$worker.WaitForExit(60000)) {
        Stop-Process -Id $worker.Id -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $pidPath) {
            $ownedWord = Get-Process -Id ([int](Get-Content -LiteralPath $pidPath)) -ErrorAction SilentlyContinue
            if ($null -ne $ownedWord -and $ownedWord.ProcessName -eq 'WINWORD' -and $ownedWord.StartTime -ge $workerStartTime.AddSeconds(-5)) { $ownedWord.Kill() }
        }
        throw ('Word timed out; see ' + $log)
    }
    Get-Content -LiteralPath $log | Out-Host
    if ($worker.ExitCode -ne 0) { throw (Get-Content -LiteralPath $errorLog -Raw) }
    if (!(Test-Path -LiteralPath $pdf) -or (Get-Item -LiteralPath $pdf).Length -eq 0) { throw 'Word produced no PDF.' }
    $wordCopy = [IO.Path]::ChangeExtension($docPath, 'word.docx')
    $before = (Invoke-Docx @('read', $docPath, '--json')) -join "`n" | ConvertFrom-Json
    $after = (Invoke-Docx @('read', $wordCopy, '--json')) -join "`n" | ConvertFrom-Json
    if ($before.Equations.Count -ne $after.Equations.Count) { throw 'Equation count changed during the Word round trip.' }
    Invoke-Docx @('validate', $wordCopy) | Out-Host
}
Write-Output ('Inspect both PDFs in ' + $OutputDirectory)
