[CmdletBinding()]
param(
    [string]$SourceDirectory = (Join-Path $PSScriptRoot '..\src'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\build\vbe-import')
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

$utf8 = [System.Text.UTF8Encoding]::new($false)
$gbk = [System.Text.Encoding]::GetEncoding(936)
$sourceFiles = @(Get-ChildItem -LiteralPath $sourcePath -File |
    Where-Object { $_.Extension.ToLowerInvariant() -in @('.bas', '.cls', '.frm') } |
    Sort-Object Name)

if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files found in $sourcePath"
}

foreach ($sourceFile in $sourceFiles) {
    $text = [System.IO.File]::ReadAllText($sourceFile.FullName, $utf8)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $text = $text -replace "`n", "`r`n"

    $destination = Join-Path $outputPath $sourceFile.Name
    [System.IO.File]::WriteAllText($destination, $text, $gbk)
    Write-Output ("Prepared {0}" -f $sourceFile.Name)
}

Get-ChildItem -LiteralPath $sourcePath -File -Filter '*.frx' -ErrorAction SilentlyContinue |
    ForEach-Object {
        $destination = Join-Path $outputPath $_.Name
        [System.IO.File]::Copy($_.FullName, $destination, $true)
        Write-Output ("Copied {0}" -f $_.Name)
    }

Write-Output "VBE import files: $outputPath"
