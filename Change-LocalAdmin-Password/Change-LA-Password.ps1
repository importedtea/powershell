<#----------------------
    Require User Choice
------------------------#>
param(
    [Parameter(Mandatory=$True)]
        [string]$User
    )

<#----------------------
    Password Management
------------------------#>
Clear-Host

Write-Host "Enter the new password for $User"

$Password = Read-Host "Enter new password" -AsSecureString
$Confirm = Read-Host "Verify your password" -AsSecureString

$P1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
$P2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Confirm))

if($P1 -ne $P2) { Write-Warning "Passwords do not match"; exit}

<#----------------------
    Get Computer List
------------------------#>
$InputFile = "computers.txt"
if(!(Test-Path $InputFile)) { Write-Warning "File $InputFile not found."; exit }

$Computers = Get-Content $InputFile

foreach ($Item in $Computers) {
    $Computer = $Item.ToUpper()
    $Connection = "OFFLINE"
    $Status = "SUCCESS"

    <#----------------------
        Check Connection
    ------------------------#>
    $Title = "`nChecking connection on $Computer"; $temp=$Title.Length; 
    Write-Host $Title
    Write-Host $("-" * $temp)
    if((Test-Connection -ComputerName $Computer -count 1 -ErrorAction 0)) { 
        $Connection="ONLINE"; Write-Host "$Computer is Online`n" -ForegroundColor Green
    } else {
        Write-Host "$Computer is OFFLINE" -ForegroundColor Red
    }

    <#---------------------------
        Attempt Password Change
    -----------------------------#>
    try {
        $Account = [ADSI]("WinNT://$Computer/$User,user")
        $Account.psbase.invoke("setpassword",$P1)
        Write-Host "Password change successful for $User`n" -ForegroundColor Green
    }
    catch {
        $Status = "FAILED"
        Write-Host "Failed to change password for $User." -ForegroundColor Red
        Write-Host "Error: $_`n" -BackgroundColor Red
    }

    <#----------------------
        Object Output
    ------------------------#>
    $Object = New-Object -TypeName PSObject -Property @{
        ComputerName = $Computer
        Connection = $Connection
        PasswordChange = $Status
    }

    $Object | Select-Object ComputerName, Connection, PasswordChange
}
