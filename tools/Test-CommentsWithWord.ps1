param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../artifacts/comment-smoke'),
    [switch]$WordWorker,
    [string]$DocumentPath,
    [string]$ExpectationPath,
    [string]$ProcessIdPath
)

$ErrorActionPreference = 'Stop'

# Only synthetic documents are opened. The STA worker is bounded by the parent;
# timeout cleanup only targets the Word process owned by this worker.
if ($WordWorker) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class CommentWordWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
    function Release-Com($value) {
        if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
        }
    }
    function Assert-WordComments($document, $expected) {
        $comments = $document.Comments
        try {
            $expectedCount = $expected.Count
            foreach ($item in $expected) { $expectedCount += $item.Replies.Count }
            if ($comments.Count -ne $expectedCount) { throw "Expected $expectedCount comments including replies; Word reports $($comments.Count)." }
            foreach ($item in $expected) {
                $matched = $false
                for ($i = 1; $i -le $comments.Count; $i++) {
                    $comment = $comments.Item($i)
                    $range = $comment.Range
                    try {
                        if ($range.Text.Trim([char[]]"`r`n`a") -ne $item.Body) { continue }
                        $matched = $true
                        $scope = $comment.Scope
                        $replies = $comment.Replies
                        try {
                            if ($scope.Text.Trim([char[]]"`r`n`a") -ne $item.Scope) { throw "Wrong Word anchor for '$($item.Body)': '$($scope.Text)'." }
                            if ([bool]$comment.Done -ne [bool]$item.Resolved) { throw "Wrong resolution state for '$($item.Body)'." }
                            if ($replies.Count -ne $item.Replies.Count) { throw "Wrong reply count for '$($item.Body)': expected $($item.Replies.Count), found $($replies.Count)." }
                            for ($j = 1; $j -le $replies.Count; $j++) {
                                $reply = $replies.Item($j)
                                $replyRange = $reply.Range
                                try {
                                    if ($replyRange.Text.Trim([char[]]"`r`n`a") -ne $item.Replies[$j - 1]) { throw 'Word reply text changed.' }
                                } finally { Release-Com $replyRange; Release-Com $reply }
                            }
                        } finally { Release-Com $replies; Release-Com $scope }
                        break
                    } finally { Release-Com $range; Release-Com $comment }
                }
                if (!$matched) { throw "Word did not retain comment '$($item.Body)'." }
            }
            Write-Output "Word verified $($expected.Count) comment anchors, reply lists, and resolution states."
        } finally { Release-Com $comments }
    }
    $word = $null
    $documents = $null
    $document = $null
    try {
        $expected = Get-Content -LiteralPath $ExpectationPath -Raw | ConvertFrom-Json
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $documents = $word.Documents
        $document = $documents.Open($DocumentPath, $false, $true, $false)
        $window = $document.ActiveWindow
        [uint32]$ownedProcessId = 0
        [void][CommentWordWindow]::GetWindowThreadProcessId([IntPtr]$window.Hwnd, [ref]$ownedProcessId)
        Set-Content -LiteralPath $ProcessIdPath -Value $ownedProcessId
        Release-Com $window
        Assert-WordComments $document $expected
        $savedPath = [IO.Path]::ChangeExtension($DocumentPath, 'word.docx')
        $document.SaveAs2([ref]$savedPath)
        $document.Close(0)
        Release-Com $document
        $document = $null
        $document = $documents.Open($savedPath, $false, $true, $false)
        Assert-WordComments $document $expected
        $window = $document.ActiveWindow
        $view = $window.View
        $view.ShowRevisionsAndComments = $true
        $view.ShowComments = $true
        Release-Com $view
        Release-Com $window
        # wdExportDocumentWithMarkup = 7 includes review balloons in the PDF.
        $pdfPath = [IO.Path]::ChangeExtension($DocumentPath, 'pdf')
        $document.ExportAsFixedFormat($pdfPath, 17, $false, 0, 0, 1, 1, 7)
        Write-Output 'Word saved, reopened, and exported a PDF with review markup.'
    } finally {
        if ($null -ne $document) { $document.Close(0); Release-Com $document }
        Release-Com $documents
        if ($null -ne $word) { $word.Quit(); Release-Com $word }
    }
    exit
}

$repoPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($OutputDirectory)
& dotnet build (Join-Path $repoPath 'src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj') -c Release -v quiet
if ($LASTEXITCODE -ne 0) { throw 'CLI build failed.' }
$cliDll = Join-Path $repoPath 'src/Lokad.DocxEdit.Cli/bin/Release/net10.0/Lokad.DocxEdit.Cli.dll'
function Invoke-Docx([string[]]$CommandArgs) {
    $result = & dotnet $cliDll @CommandArgs
    if ($LASTEXITCODE -ne 0) { throw ('docxedit failed: ' + ($CommandArgs -join ' ')) }
    return $result
}
function Apply-Patch([string]$inputPath, [string]$name, [string]$patch) {
    $patchPath = Join-Path $OutputDirectory ($name + '.docxpatch')
    $outputPath = Join-Path $OutputDirectory ($name + '.docx')
    Set-Content -LiteralPath $patchPath -Value $patch -Encoding utf8
    Invoke-Docx @('check', $inputPath, $patchPath) | Out-Host
    Invoke-Docx @('apply', $inputPath, $patchPath, '--output', $outputPath) | Out-Host
    Invoke-Docx @('validate', $outputPath) | Out-Host
    return $outputPath
}
function Test-WithWord([string]$path, [object[]]$expectations) {
    $expectPath = [IO.Path]::ChangeExtension($path, 'expected.json')
    ConvertTo-Json -InputObject $expectations -Depth 5 | Set-Content -LiteralPath $expectPath -Encoding utf8
    $log = [IO.Path]::ChangeExtension($path, 'word.log')
    $errorLog = [IO.Path]::ChangeExtension($path, 'word-error.log')
    $pidPath = [IO.Path]::ChangeExtension($path, 'word.pid')
    if (Test-Path -LiteralPath $pidPath) { Remove-Item -LiteralPath $pidPath }
    $arguments = @('-NoProfile', '-STA', '-File', ('"' + $PSCommandPath + '"'), '-WordWorker',
        '-DocumentPath', ('"' + $path + '"'), '-ExpectationPath', ('"' + $expectPath + '"'), '-ProcessIdPath', ('"' + $pidPath + '"'))
    $worker = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $log -RedirectStandardError $errorLog
    # Retain the handle so Windows PowerShell can read ExitCode after the wait.
    $null = $worker.Handle
    $started = $worker.StartTime
    if (!$worker.WaitForExit(60000)) {
        Stop-Process -Id $worker.Id -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $pidPath) {
            $ownedWord = Get-Process -Id ([int](Get-Content -LiteralPath $pidPath)) -ErrorAction SilentlyContinue
            if ($null -ne $ownedWord -and $ownedWord.ProcessName -eq 'WINWORD' -and $ownedWord.StartTime -ge $started.AddSeconds(-5)) { $ownedWord.Kill() }
        }
        throw ('Word timed out; see ' + $log)
    }
    Get-Content -LiteralPath $log | Out-Host
    if ($worker.ExitCode -ne 0) { throw (Get-Content -LiteralPath $errorLog -Raw) }
    $pdf = [IO.Path]::ChangeExtension($path, 'pdf')
    if (!(Test-Path -LiteralPath $pdf) -or (Get-Item -LiteralPath $pdf).Length -eq 0) { throw 'Word produced no PDF.' }
    $copy = [IO.Path]::ChangeExtension($path, 'word.docx')
    Invoke-Docx @('validate', $copy) | Out-Host
    return $copy
}

$blank = Join-Path $OutputDirectory 'blank.docx'
Invoke-Docx @('create', '--output', $blank) | Out-Host
# Seed a synthetic table and a bookmark to exercise existing document markup.
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$zip = [IO.Compression.ZipFile]::Open($blank, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $zip.GetEntry('word/document.xml')
    $reader = New-Object IO.StreamReader($entry.Open())
    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $body = $xml.SelectSingleNode('//w:body', $ns)
    $body.InnerXml = @'
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Review comments: beta needs evidence.</w:t></w:r></w:p>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bookmarkStart w:id="8" w:name="fixed"/><w:r><w:t xml:space="preserve">Bookmarked text. </w:t></w:r><w:bookmarkEnd w:id="8"/><w:r><w:t>Reviewed paragraph.</w:t></w:r></w:p>
<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblPr><w:tblBorders><w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>First estimate.</w:t></w:r></w:p><w:p><w:r><w:t>Second estimate.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
'@
    $stream = $entry.Open()
    try { $stream.SetLength(0); $xml.Save($stream) } finally { $stream.Dispose() }
} finally { $zip.Dispose() }

$generated = Apply-Patch $blank 'comments-generated' @'
docxpatch 1
op add-comment
target M.P0001
anchor-text beta
text Please clarify beta.
author Reviewer
as review
end
op add-comment-reply
target @review
text Clarification is pending.
author Author
end
op add-comment
target M.P0002
anchor-text Reviewed paragraph.
text This paragraph has been reviewed.
as closed
end
op resolve-comment
target @closed
end
op add-comment
target M.T0001.R01.C01
anchor-text estimate
occurrence 2
text Explain the second estimate.
end
op add-comment
target M.P0001
anchor-text evidence
text Add a source.
end
'@
$expected = @(
    @{ Body = 'Please clarify beta.'; Scope = 'beta'; Resolved = $false; Replies = @('Clarification is pending.') },
    @{ Body = 'This paragraph has been reviewed.'; Scope = 'Reviewed paragraph.'; Resolved = $true; Replies = @() },
    @{ Body = 'Explain the second estimate.'; Scope = 'estimate'; Resolved = $false; Replies = @() },
    @{ Body = 'Add a source.'; Scope = 'evidence'; Resolved = $false; Replies = @() }
)
$wordCopy = Test-WithWord $generated $expected
$changes = (Invoke-Docx @('changes', $wordCopy, '--include-comment-text', '--json')) -join "`n" | ConvertFrom-Json
function Comment-Id([string]$body) {
    $matches = @($changes.CommentSummary | Where-Object { $_.TextSnippet -eq $body })
    if ($matches.Count -ne 1) { throw "Comment lookup is ambiguous or missing: $body" }
    return ('comment:' + $matches[0].CommentId)
}
$rootId = Comment-Id 'Please clarify beta.'
$replyId = Comment-Id 'Clarification is pending.'
$closedId = Comment-Id 'This paragraph has been reviewed.'
$sourceId = Comment-Id 'Add a source.'
$edited = Apply-Patch $wordCopy 'comments-edited' @"
docxpatch 1
op set-comment-text
target $rootId
text Please clarify beta with a definition.
end
op set-comment-text
target $replyId
text Definition added in the next revision.
end
op add-comment-reply
target $rootId
text The definition is now clear.
author Reviewer
end
op reopen-comment
target $closedId
end
op delete-comment
target $sourceId
end
"@
$expected = @(
    @{ Body = 'Please clarify beta with a definition.'; Scope = 'beta'; Resolved = $false; Replies = @('Definition added in the next revision.', 'The definition is now clear.') },
    @{ Body = 'This paragraph has been reviewed.'; Scope = 'Reviewed paragraph.'; Resolved = $false; Replies = @() },
    @{ Body = 'Explain the second estimate.'; Scope = 'estimate'; Resolved = $false; Replies = @() }
)
$editedCopy = Test-WithWord $edited $expected
$changes = (Invoke-Docx @('changes', $editedCopy, '--include-comment-text', '--json')) -join "`n" | ConvertFrom-Json
$rootId = Comment-Id 'Please clarify beta with a definition.'
$replyId = Comment-Id 'Definition added in the next revision.'
$secondReplyId = Comment-Id 'The definition is now clear.'
$deleted = Apply-Patch $editedCopy 'comments-deleted' @"
docxpatch 1
op delete-comment-reply
target $replyId
end
op delete-comment-reply
target $secondReplyId
end
op delete-comment
target $rootId
end
"@
Test-WithWord $deleted @($expected[1], $expected[2]) | Out-Host
Write-Output ('Inspect the three review PDFs in ' + $OutputDirectory)
