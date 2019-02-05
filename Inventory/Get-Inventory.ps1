<#-------------------------
    HTML Header & Style
--------------------------#>
$Header = @"
<center>
    <style>
        body { 
            background-color:#E5E4E2;
            font-family:Times New Roman;
            font-size:12pt;
            position:relative;
        }
        th{ 
            border-width: 1px; padding: 0px; border-style: solid; border-color: black; background-color: #6495ED;
        }
        td {
            border-width: 1px; padding: 0px; border-style: solid; border-color: black; width: 1px; white-space: nowrap;
            height: 5px;
            text-align: center;
        }
        table, tr, td, th { padding: 3px; margin: 0px ;white-space:pre; }
        table { 
            margin-left:auto;
            margin-right:auto;
            border-width: 2px; border-style: solid; border-color: black; border-collapse: collapse;
        }
        h2 {
            text-align: center;
            font-family:Tahoma;
            color:#6D7B8D;
        }

        th[scope="col"] {
            background: #ccc;
        }
        tr:nth-child(even) {
            background: #eee;
        }
        tr:nth-child(odd) {
            background: #ddd;
        }

        #footer {
            color: black;
            text-align:right;
            position: static;
            font-size: small;
          }
    </style>
    <title>Computer Information Report</title>
</center>
"@

<#-------------------------
    Menu Function
--------------------------#>
function Show-Menu {
    Param (
        [Parameter(Mandatory=$false, Position=0)]
        [string] $Title
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Demo"
    Write-Host "Q: Press 'Q' to quit."
    Write-Host "================ $Title ================`n"
}

<#-------------------------
    Generate Switch
--------------------------#>
function Switch-Produce () {
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Key,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Class,
        [Parameter(Mandatory=$true, Position=2)]
        [string] $ClassPC
    )
    $global:key = $key
    $global:Class = $Class
    $global:ClassPC = "${ClassPC}-computers.txt"

    Write-Host "You chose $key" -ForegroundColor Yellow
    Write-Host "The class directory is: $Class" -ForegroundColor Yellow
    Write-Host "The computers list is: ${ClassPC}-computers.txt" -ForegroundColor Yellow
}

<#-------------------------
    Error Preference
--------------------------#>
$ErrorActionPreference = 'Stop'

<#-------------------------
    Start Execution
--------------------------#>
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

<#-------------------------
        Menu Loop
--------------------------#>
do {
    Show-Menu -Title "Classrooms"
    $Choice = Read-Host "What classroom would you like to inventory?"

    switch ($Choice) {
        1 { Switch-Produce "Demo" demo demo }
        'q' { exit }
        default { Show-Menu }
    }
} while ($Choice -eq 'q')

<#-------------------------
    Array Creation
--------------------------#>
$Results = @()

<#-------------------------
        Computers
--------------------------#>
$Computers = (Get-Content $PSScriptRoot\Classrooms\$Class\$ClassPC)

<#-------------------------
    Directory/File Creation
--------------------------#>
$Filename = Read-Host "What would you like your filename to be? "
$Dirname = New-Item -Path "$PSScriptRoot\Classrooms\$Class\" -Name "inv" -ItemType Directory -Force

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
            Domain = (Get-WmiObject -ComputerName $computer -Class win32_computersystem).Domain
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
        $Properties.Age = (New-TimeSpan -Start ($Properties.Age.ConvertToDateTime($Properties.Age.releasedate).ToShortDateString()) -End $(Get-Date)).Days / 365
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
$Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | Export-Csv -Path $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv -NoTypeInformation

<#-------------------------
        HTML Export
--------------------------#>
$PCTitle = "<h2>Computer Information Report for $Key</h2>"

## Create the HTML file
$Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | ConvertTo-Html -Head $Header -Body $PCTitle | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html

## Append a footer
$footer = "<footer id='footer'><i>$(get-date -Format g)</i></footer>"
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
Write-Host "The script took $Time seconds`n"