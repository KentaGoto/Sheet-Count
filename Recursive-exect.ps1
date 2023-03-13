Param (
    [string]$FolderPath
)

$FolderPath = Read-Host -Prompt "Dir"
$scriptPath = "$PSScriptRoot\Sheet-Count.py"

Get-ChildItem -Path $FolderPath -Recurse | ForEach-Object { 
    if (-not $_.PSIsContainer -and $_.Extension -eq ".xlsx") {
        Start-Process -FilePath "python.exe" -ArgumentList $scriptPath, $_.FullName
    }
}
