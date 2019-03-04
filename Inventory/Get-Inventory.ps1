<#
.SYNOPSIS  
Gets inventory of computers listed in various locations.

.DESCRIPTION
Generate quick and concise inventory with little manual maintenance. 

.PARAMETER Simple
Reduces the output for those who don't care about specifics.

.EXAMPLE
Get-Inventory.ps1

.EXAMPLE
Get-Inventory.ps1 -Simple

.NOTES
Author: Matthew Fabrizio
Version: 2.0.0

________                                 _____       
___  __ \____  ______________ _______ ______(_)______
__  / / /_  / / /_  __ \  __ `/_  __ `__ \_  /_  ___/
_  /_/ /_  /_/ /_  / / / /_/ /_  / / / / /  / / /__  
/_____/ _\__, / /_/ /_/\__,_/ /_/ /_/ /_//_/  \___/  
        /____/                                       

Future Implementation:
Get a list of computers from AD (priority = -1)

.LINK
    https://www.github.com/importedtea/powershell/Remote_Tools/
#>

param (
    [Parameter(Mandatory=$false)]
    [switch] $Simple
)

function Show-Menu {
    Clear-Host

    <#-------------------------
        Print Menu Title
    --------------------------#>
    Write-Host "************ MENU ************"

    <#-------------------------
        Print Menu Entries
    --------------------------#>
    $Entry = 1
    foreach ($Item in $List) {
        Write-Host "[$Entry]: $($Item)"
        $Entry++
    }
  
    <#-------------------------
        Add Static Entries
    --------------------------#>
    Write-Host "[Q]: Quit`n"
      
    <#-------------------------
        Accept User Input
    --------------------------#>
    $Selection = Read-Host "Select a menu option"
  
    <#-------------------------
        Validate STDIN
    --------------------------#>
    if ($Selection -match "^\d+$") {
        # If one menu item, read that item
        if ((Get-ChildItem "$PSScriptRoot\menu-entries.csv").Name.Count -eq 1) {
            $Choice = $List
        }
        # If menu entry >=2
        else {
            # If Selection is numeric
            $Choice = $($List[$Selection - 1])
        }
    }
    else {
        # If Selection is alpha
        $Choice = $Selection
    }

    <#-------------------------
        Return the Choice
    --------------------------#>
    return $Choice;
}

function Construct_Iventory() {
    <#-------------------------
        Create Root Directory
    --------------------------#>
    if (!(Test-Path "$PSScriptRoot\Classrooms\")) {
        New-Item -Path . -ItemType Directory -Name "Classrooms" | Out-Null
    }
      
    <#-------------------------
        Create Menu Loop
    --------------------------#>
    foreach ($Classroom in $List) {
        <#-------------------------
            Initialize Helper Vars
        --------------------------#>
        $TEMP_DIR = $Classroom.ToLower().replace(" ", "_")
        $TEMP_FILE = $Classroom.ToLower().replace(" ", "-") + '-computers' + '.txt'
        $FULL_PATH = "$PSScriptRoot\Classrooms\$TEMP_DIR\$TEMP_FILE"
          
        <#-------------------------
            Create Root\Directories
        --------------------------#>
        New-Item -Path "$PSScriptRoot\Classrooms\" -ItemType Directory -Name $TEMP_DIR -Force | Out-Null
  
        <#-------------------------
            Generate computer.txt files
        --------------------------#>
        if (!(Test-Path $FULL_PATH)) {
            $TEMP_RETVAL = $true  
            New-Item -Path "$PSScriptRoot\Classrooms\$TEMP_DIR\" -ItemType File -Name $TEMP_FILE | Out-Null
            Write-Host "You have added a new menu entry called $TEMP_FILE, please populate it now" -ForegroundColor Yellow; continue
        }

        <#-------------------------
            Validate Content
        --------------------------#>
        if ($Null -eq (Get-Content $FULL_PATH)) {
            $TEMP_RETVAL = $true
            Write-Host "You are missing content for $TEMP_FILE, please populate it now." -ForegroundColor Yellow; continue
        }
    }
    <#-------------------------
        If NULL data; exit
    --------------------------#>
    if ($TEMP_RETVAL) { exit }
}

<#-------------------------
    Grab all menu entries
--------------------------#>
if (!(Test-Path "$PSScriptRoot\menu-entries.csv")) { 
    New-Item -Path "$PSScriptRoot\" -ItemType File -Name "menu-entries.csv" | Out-Null
    Write-Host "You had no menu entries, please populate menu-entries.csv now" -ForegroundColor Yellow; exit
}
elseif ($Null -eq (Get-Content "$PSScriptRoot\menu-entries.csv")) {
    Write-Host "You are missing content for menu-entries.csv, please populate it now." -ForegroundColor Yellow; exit
}
else { $List = Get-Content "$PSScriptRoot\menu-entries.csv" }
  
<#-------------------------
    Grab HTML Styling
--------------------------#>
$Header = Get-Content "$PSScriptRoot\assets\style.html"

<#-------------------------
    Call Construct_Iventory
    Generate/Validate data
--------------------------#>
Construct_Iventory

<#-------------------------
    Menu Selection Logic
--------------------------#>
do {
    $Choice = Show-Menu

    switch ($Choice) {
        { $List -contains $Choice } {
            $Key = $Choice
            $Shop = $Choice.ToLower().replace(" ", "-")
            $ShopPC = "$Choice-computers.txt".ToLower().Replace(" ", "-")

            # Match if $list array contains the $selected $Choice object
            Write-Host "`nYou chose $($Choice)." -ForegroundColor Green
            Write-Host "Inventory will be pulled from $Shop\$ShopPC`n" -ForegroundColor Green
        }
        'q' { exit }
        '\' { Write-Host "dev setting" -ForegroundColor Green }
        default { Write-Host "Invalid menu choice" -ForegroundColor Red }  
    }
} while ($Choice -eq 'q')

<#-------------------------
    Error Preference
--------------------------#>
$ErrorActionPreference = 'Stop'

<#-------------------------
    Start Execution
--------------------------#>
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

<#-------------------------
    Array Creation
--------------------------#>
$Results = @()

<#-------------------------
        Computers
--------------------------#>
$Computers = (Get-Content "$PSScriptRoot\Classrooms\$Shop\$ShopPC")

<#-------------------------
    Directory/File Creation
--------------------------#>
$Filename = Read-Host "What would you like your filename to be? "
$Dirname = New-Item -Path "$PSScriptRoot\Classrooms\$Shop\" -Name "inv" -ItemType Directory -Force

<#-------------------------
    $Computer Loop
--------------------------#>
ForEach ($computer in $Computers) {
    try {
        <#-------------------------
            Properties Object
        --------------------------#>
        $Properties = @{
            Hostname = (Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem).CSName
            Manufacturer = (Get-WmiObject -ComputerName $computer -Class Win32_ComputerSystem).Manufacturer
            Model = (Get-WmiObject -ComputerName $computer -Class Win32_ComputerSystem).Model
            Serial = (Get-WmiObject -ComputerName $computer -Class Win32_Bios).SerialNumber
            Edition = (Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem).Caption
            OS = (Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem).Version
            Memory = Get-WmiObject -ComputerName $computer -Class Win32_ComputerSystem | Select-Object TotalPhysicalMemory, @{name="GB";expr={[float]($_.TotalPhysicalMemory/1GB)}} | Select-Object -ExpandProperty GB
            IP = ([System.Net.Dns]::GetHostByName("$computer").AddressList[0]).IpAddressToString
            Domain = (Get-WmiObject -ComputerName $computer -Class Win32_Computersystem).Domain
            MAC = (Get-WmiObject -ComputerName $computer -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'").MACAddress | Select-Object -First 1
            Age = Get-WmiObject -ComputerName $computer -Class Win32_BIOS
            ReimageDate = Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem
        }

        <#-------------------------
            Release Conversion
        --------------------------#>
        switch($Properties.OS){
            '10.0.10240' {$Properties.OS="1507"}
            '10.0.10586' {$Properties.OS="1511"}
            '10.0.14393' {$Properties.OS="1607"}
            '10.0.15063' {$Properties.OS="1703"}
            '10.0.16299' {$Properties.OS="1709"}
            '10.0.17134' {$Properties.OS="1803"}
            '10.0.17763' {$Properties.OS="1809"}
        }

        <#-------------------------
                Calculations
        --------------------------#>
        $Properties.Age = (New-TimeSpan -Start ($Properties.Age.ConvertToDateTime($Properties.Age.ReleaseDate).ToShortDateString()) -End $(Get-Date)).Days / 365
        $Properties.ReimageDate = ($Properties.ReimageDate.ConvertToDateTime($Properties.ReimageDate.InstallDate).ToString("MM-dd-yyyy"))
        
        <#-------------------------
            Spicy Console Log
        --------------------------#>
        [PSCustomObject]@{
            Hostname = $Properties.Hostname
            Manufacturer = $Properties.Manufacturer
            Model = $Properties.Model
            SerialNumber = $Properties.Serial
            OSEdition = $Properties.Edition
            OS = $Properties.OS
            RAM = $Properties.Memory
            IP = $Properties.IP
            MAC = $Properties.MAC
            Domain = $Properties.Domain
            Age = $Properties.Age
            ReimageDate = $Properties.ReimageDate
        } | Format-List

        <#----------------------------------
            Append Object Data to $Results
        ------------------------------------#>
        $Results += New-Object PSObject -Property $Properties
    }
    <#-------------------------
        Error Handling
    --------------------------#>
    catch {
        Write-Host "`n$computer Offline" -ForegroundColor Red
        continue
    }
}

<#-------------------------
        CSV Export
--------------------------#>
if ($Simple) {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS | Export-Csv -Path $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv -NoTypeInformation
}
else {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | Export-Csv -Path $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv -NoTypeInformation
}

<#-------------------------
        HTML Export
--------------------------#>
# $PCTitle = "<h2>Computer Information Report for $Key</h2>"
$PCTitle = "<div class='header'>Computer Information Report for $Key</div>"

# Create the HTML file
if ($Simple) {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS | ConvertTo-Html -Head $Header -Body $PCTitle | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html

}
else {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | ConvertTo-Html -Head $Header -Body $PCTitle | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html
 
}

## Append a footer
# $footer = "<footer id='footer'><i>$(get-date -Format g)</i></footer>"
$footer = "<div class='Copyright'>$((Get-Date -Format g))</div>"
$footer | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html -Append

<#-------------------------
    Report Locations
--------------------------#>
Write-Host "You can view the logfile at $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv" -ForegroundColor Green
Write-Host "You can view the HTML report at $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv" -ForegroundColor Green

<#-------------------------
    Script Execution
--------------------------#>
$Stopwatch.Stop()
$Time = $Stopwatch.Elapsed
Write-Host "`nThe script took $Time seconds`n"