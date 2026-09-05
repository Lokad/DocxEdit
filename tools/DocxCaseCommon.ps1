# Shared helpers for the DocxEdit PowerShell tooling.
# Dot-sourced by CheckPrivateCase.ps1 and RunAgentChallenge.ps1.
# Assert-PrivatePath requires the caller to define $PrivateRoot; the other helpers need no globals.
# RunAgentChallenge.ps1 keeps its own Invoke-ProcessCapture (timeouts and stdin).

function Test-IsUnderPath([string] $Path, [string] $Root) {
    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $resolvedRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $resolvedPath.Equals($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $resolvedPath.StartsWith($resolvedRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RepoRelativePath([string] $Path) {
    $root = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $rootUri = [System.Uri]::new($root)
    $pathUri = [System.Uri]::new($fullPath)
    return [System.Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('\', '/')
}

function Assert-PrivatePath([string] $Path, [string] $Kind) {
    if (-not (Test-IsUnderPath $Path $PrivateRoot)) {
        throw "$Kind must live under private-cases/."
    }

    $relative = Get-RepoRelativePath $Path
    $tracked = & git -C $RepoRoot ls-files -- $relative
    if ($tracked) {
        throw "$Kind must not be tracked by git."
    }

    & git -C $RepoRoot check-ignore -q -- $relative
    if ($LASTEXITCODE -ne 0) {
        throw "$Kind must be ignored by git."
    }
}

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

function Get-FileHashSha256([string] $Path) {
    if ($null -eq (Get-Command Get-FileHash -ErrorAction SilentlyContinue)) {
        throw "File hashing needs the Get-FileHash cmdlet (pwsh 7+ or a full Windows PowerShell 5.1 install); refusing to continue with an unverified copy."
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
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

