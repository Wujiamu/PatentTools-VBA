[CmdletBinding()]
param(
    [string]$SourceDirectory = (Join-Path $PSScriptRoot '..\src'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\build\vbe-import'),
    [switch]$AsciiFileNames
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sourcePath = (Resolve-Path -LiteralPath $SourceDirectory).Path
$repositoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
} else {
    $outputPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
}

if (-not $outputPath.StartsWith($repositoryPath + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Output directory must remain inside the repository: $outputPath"
}

New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

$utf8 = [System.Text.UTF8Encoding]::new($false, $true)
$gbk = [System.Text.Encoding]::GetEncoding(
    936,
    [System.Text.EncoderExceptionFallback]::new(),
    [System.Text.DecoderExceptionFallback]::new()
)
$sourceFiles = @(Get-ChildItem -LiteralPath $sourcePath -File |
    Where-Object { $_.Extension.ToLowerInvariant() -in @('.bas', '.cls', '.frm') } |
    Sort-Object Name)

if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files found in $sourcePath"
}

foreach ($sourceFile in $sourceFiles) {
    $sourceBytes = [System.IO.File]::ReadAllBytes($sourceFile.FullName)
    $text = $utf8.GetString($sourceBytes)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $text = $text -replace "`n", "`r`n"

    # 先编码到内存并严格回读，避免将中文静默替换成问号。
    $encoded = $gbk.GetBytes($text)
    $roundTrip = $gbk.GetString($encoded)
    if ($roundTrip -cne $text) {
        throw "GBK round-trip mismatch: $($sourceFile.Name)"
    }

    $destinationName = $sourceFile.Name
    if ($AsciiFileNames) {
        $moduleName = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"').Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($moduleName)) {
            $moduleName = $sourceFile.BaseName
        }
        $safeName = ($moduleName -replace '[^A-Za-z0-9_.-]', '_')
        $destinationName = "$safeName$($sourceFile.Extension)"
    }

    $destination = Join-Path $outputPath $destinationName
    [System.IO.File]::WriteAllBytes($destination, $encoded)
    Write-Output ("Prepared {0} -> {1}" -f $sourceFile.Name, $destinationName)
}

Get-ChildItem -LiteralPath $sourcePath -File -Filter '*.frx' -ErrorAction SilentlyContinue |
    ForEach-Object {
        $destination = Join-Path $outputPath $_.Name
        [System.IO.File]::Copy($_.FullName, $destination, $true)
        Write-Output ("Copied {0}" -f $_.Name)
    }

Write-Output "VBE import files: $outputPath"
