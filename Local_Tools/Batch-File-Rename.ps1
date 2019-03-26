<#----------------------------------
            begin
------------------------------------#>
$Backup_DIR = "$PSScriptRoot\PIC_BACKUP"
$IMG_DIR = "$PSScriptRoot\IMAGES"

$Files = Get-ChildItem "$PSScriptRoot\IMAGES\"
$CSV = Import-Csv "$PSScriptRoot\student-photo-list.csv"
$Line = 0

<#----------------------------------
    check if any backup images exist
------------------------------------#>
if (!(Test-Path $Backup_DIR)) { New-Item -ItemType Directory -Force -Path $Backup_DIR | Out-Null }

$Backup_Files = Get-ChildItem "$Backup_DIR\"

if ($Backup_Files.Count -eq 0) { 
    Write-Host "Directory is empty...backing up pictures" -ForegroundColor Yellow
    Copy-Item -Path "$IMG_DIR\*" -Destination "$Backup_DIR\" -Verbose
}

if ($Backup_Files.Count -eq $Files.Count) {
    Write-Host "$Backup_DIR and $IMG_DIR have matching content length" -ForegroundColor Green
}

<#----------------------------------
    check if script has been executed
------------------------------------#>
if ($Files.Name -like "000*") { Write-Host "File samples matche 000*, check contents of $IMG_DIR" -ForegroundColor Yellow; exit}

<#----------------------------------
    loop through all filenames; change names to ID
------------------------------------#>
foreach ($Filename in $Files) {
    $ID = $CSV[$Line].ID

    Rename-Item "$PSScriptRoot\IMAGES\$Filename" -NewName "$ID.JPG" -Force

    if ($LASTEXITCODE -eq 0) { Write-Host "`nRenaming filename $Filename to $ID.JPG" -ForegroundColor Green }
    else { Write-Error "File: $Filename failed to rename to $ID.jpg" }
    
    $Line++
}