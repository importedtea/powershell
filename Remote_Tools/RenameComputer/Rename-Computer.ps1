[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [string] $Domain,

    [Parameter(Mandatory)]
    [string] $Username
)

$CSV = Import-Csv "names.csv"

$Cred = Get-Credential -UserName $Domain\$Username -Message "Enter your stinky password, jak"

foreach ($Line in $CSV) {
    try {
        Rename-Computer -ComputerName $Line.OldName -NewName $Line.NewName -DomainCredential $Cred -Force -PassThru -Restart -ErrorAction Stop -WhatIf
        Write-Host "$($Line.OldName) renamed successfully to $($Line.NewName)" -ForegroundColor Green
    }
    catch {
        Write-Error "$($Line.OldName) cannot be reached. Check to see if it is online, then try again."
    }
}