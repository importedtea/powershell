<#
.SYNOPSIS  
Produce litter. Easily access your stinky clumps.

.DESCRIPTION
Produce little clumps of inventory stored in a soft, delicate bed of dust.
Inspired by my girlfriend, as she quotes:
    
    "
    Well it is where your inventory is stored which could include many things
    just like how monkey stores lots of her inventory in her LitterBox
    "

.PARAMETER Simple
Reduces the output for stinky doggos.

.EXAMPLE
Litter-Box.ps1

.EXAMPLE
Litter-Box.ps1 -Simple

.NOTES
Author  : Matthew Fabrizio
Project : Litter-Box
Version : 2
                                     
              ,-.       _,---._
             /  )    .-'       `.     
            (  (   ,'            `  ___ 
             \  `-"             \'    / |
              `.              ,  \   /  |
               /`.          ,'-`----Y   |
              (            ;        |   '
              |  ,-.    ,-'         |  /
              |  | (   | litter-box | /
              )  |  \  `.___________|/
              `--'   `--'


Version 1.0.0
A very static menu-based inventory system was introduced.

Version 2.0.0
Dynamic Menu added! 
                                    
Version 2.1.1
Non-Discoverable implementation was added.                                                                                                            
Squashed some bugs.

Version 2.2.1
Script name changed to Litter-Box.ps1
Added Add-Content functionality for on the fly inventorying. Access with a|A in the menu.
Added compatibility for a company logo at the top of the screen.
Add a help function and paramter -Help

Version 2.3.1
Added a variable for Classrooms for easy change of directory name;
Condensed static variables into one location for ease of use; see Begin comment
Added a history file for each newly created inventory domain
Added a Development flag to enable debugging
Simplified inventory creation process (i.e. turned three steps into two for new items)
Sorted the menu by alpha
Added HTMLOut and CSVOut variables because of a giant mess
Provided an option (b|B) to run on a single hostname
Added a todo.csv file for notes about an inventory domain
Added non-discoverable report scanning with option (n|N)

Future Implementation:
If an age field is given in a csv, use that to calculate based on current time/date. (ex. printers.csv has Age column that says 0 (i.e. new), calculate against Get-Date to verify age when script was executed)
Look into color coding specific HTML rows (i.e. if the date is > 7 highlight red) - probably will only work for non-disc

.LINK
    https://www.github.com/importedtea/powershell/Remote_Tools/
#>

param (
    [Parameter(Mandatory=$false)]
    [switch] $Simple,
    [switch] $Help,
    [switch] $Development
)

function Help() {
    Write-Host @"
        
        Running the script:
            .\Litter-Box
        
        Running simplified version:
            .\Litter-Box -Simple

        Launching Help:
            .\Litter-Box -Help

        Launching Debug:
            .\Litter-Box -Development
        
        Menu-Entries.csv
            menu-entries.csv is added on initial run after clone.
            Script will notify you that the file was added.
            If you try running it again, it will tell you it is empty.

            It is more preferable to add additional content with the alpha a|A option

        After Content is Added:
            When you run the script, it will notify you that you added something 
                and prompt you to add hostnames.
            Step 1: Add a hostname to *-computers.txt file.
            Optional: Add non-discoverable information in appropriate .csv file.
            Step 2: .\Litter-Box.ps1

        Choosing Your Content:
            If you followed all the warning prompts, you should successfully see a menu
                with your item and four additional alpha values.
            Step 1: Choose your menu entry
            Step 2: Give your inventory output file a name

        Adding Content:
            You have two choices of adding content. You can either use the alpha a|A option in the menu
                or manually update the menu-entries.csv file.
        
            Either option supports lower|uppercase. It is possible to break this so try to avoid special chars.

        Debugging Content:
            If you use the -Development parameter, a file called debug.log will be created in $PSScriptRoot

"@
}

function GCLN { $MyInvocation.ScriptLineNumber } 

function Write-Log ($Message, $Path="$PSScriptRoot\debug.log") { "[$(Get-Date -Format g)] => $Message" | Tee-Object -FilePath $Path -Append | Write-Verbose }

function Remove_WhiteSpace($File) {
    # Grab file contents; remove any trailing whitespace
    $Newtext = (Get-Content -Path $File -Raw) -replace "(?s)`r`n\s*$"
    [System.IO.File]::WriteAllText($File,$Newtext)
}

function Local_Scan($TempKey) {
    $Properties = @{
        Hostname = (Get-WmiObject -ComputerName $TempKey -Class Win32_OperatingSystem).CSName
        Manufacturer = (Get-WmiObject -ComputerName $TempKey -Class Win32_ComputerSystem).Manufacturer
        Model = (Get-WmiObject -ComputerName $TempKey -Class Win32_ComputerSystem).Model
        Serial = (Get-WmiObject -ComputerName $TempKey -Class Win32_Bios).SerialNumber
        Edition = (Get-WmiObject -ComputerName $TempKey -Class Win32_OperatingSystem).Caption
        OS = (Get-WmiObject -ComputerName $TempKey -Class Win32_OperatingSystem).Version
        Memory = Get-WmiObject -ComputerName $TempKey -Class Win32_ComputerSystem | Select-Object TotalPhysicalMemory, @{name="GB";expr={[float]($_.TotalPhysicalMemory/1GB)}} | Select-Object -ExpandProperty GB
        IP = ([System.Net.Dns]::GetHostByName("$TempKey").AddressList[0]).IpAddressToString
        Domain = (Get-WmiObject -ComputerName $TempKey -Class Win32_Computersystem).Domain
        MAC = (Get-WmiObject -ComputerName $TempKey -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'").MACAddress | Select-Object -First 1
        Age = Get-WmiObject -ComputerName $TempKey -Class Win32_BIOS
        ReimageDate = Get-WmiObject -ComputerName $TempKey -Class Win32_OperatingSystem
    }

    switch($Properties.OS){
        '10.0.10240' {$Properties.OS="1507"}
        '10.0.10586' {$Properties.OS="1511"}
        '10.0.14393' {$Properties.OS="1607"}
        '10.0.15063' {$Properties.OS="1703"}
        '10.0.16299' {$Properties.OS="1709"}
        '10.0.17134' {$Properties.OS="1803"}
        '10.0.17763' {$Properties.OS="1809"}
    }

    $Properties.Age = (New-TimeSpan -Start ($Properties.Age.ConvertToDateTime($Properties.Age.ReleaseDate).ToShortDateString()) -End $(Get-Date)).Days / 365
    $Properties.ReimageDate = ($Properties.ReimageDate.ConvertToDateTime($Properties.ReimageDate.InstallDate).ToString("MM-dd-yyyy"))

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
}

function NonDiscoverable_Scan($Option) {
    # Strip the extension from option; capitalize option
    $FileName = [IO.Path]::GetFileNameWithoutExtension($Option); $FileName = (Get-Culture).TextInfo.ToTitleCase("$FileName".ToLower())
    
    # Check for Reports dir; create if not found
    if (!(Test-Path "$PSScriptRoot\Reports")) { New-Item -Path "$PSScriptRoot\" -Name "Reports" -ItemType Directory -Force | Out-Null }
    
    # Define the report location; add styling to html output
    $ReportLoc = "$PSScriptRoot\Reports\$FileName.html"
    $HTMLHeader | Out-File $ReportLoc

    # Choose between sort options
    if ($FileName -eq "todo","history") { $Property = "Date" }
    else { $Property = "Asset" }

    # Loop through all files that match option
    foreach ($Item in $NonDiscoverablesFile -match "$Option") {
        # Loop through the chosen file in each inventory domain
        foreach ($File in $InvFiles.Name) {
            $Csv = Import-Csv "$PSScriptRoot\$InvDir\$File\$Item"
            
            if ($Csv.length -eq 0) { continue }
            else {
                # (1) Grab the first inventory domain title and capitalize
                # (2) Add the appropriate header with title
                # (3) Convert the CSV to HTML
                $Title = (Get-Culture).TextInfo.ToTitleCase("$File".ToLower()).Replace("_", " ")
                "<br><br><div class='Header'>$FileName Report for $Title</div>" | Out-File $ReportLoc -Append
                $Csv | Select-Object * | Sort-Object -Property $Property | ConvertTo-Html -Fragment | Out-File $ReportLoc -Append
            }
        }
    }

    Write-Host "`nYou can view the report for $FileName @ $ReportLoc`n" -ForegroundColor Green
}

function Show-Menu {
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Entered Show-Menu function" }
    if (!$Development) { Clear-Host }

    <#-------------------------
        Print Menu Title
    --------------------------#>
    Write-Host "`n************ MENU ************"
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Wrote Title" }

    <#-------------------------
        Print Menu Entries
    --------------------------#>
    $Entry = 1
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Initialized menu numbers" }
    foreach ($Item in $List) {
        Write-Host "[$Entry]: $($Item)"
        $Entry++
    }
  
    <#-------------------------
        Add Static Entries
    --------------------------#>
    Write-Host "`n[A]: Add Menu Entry"
    Write-Host "[B]: Quick Scan"
    Write-Host "[N]: Non-Discoverable Report"
    Write-Host "[Q]: Quit`n"
    
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Appended add option" }
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Appended quit option" }
      
    <#-------------------------
        Accept User Input
    --------------------------#>
    $Selection = Read-Host "Select a menu option"
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Prompted user for input; user chose $Selection" }
  
    <#-------------------------
        Validate STDIN
    --------------------------#>
    # If Selection is numeric
    if ($Selection -match "^\d+$") {
        # If one menu item, read that item
        if ((Get-Content "$PSScriptRoot\menu-entries.csv").Count -eq 1) {
            $Choice = $List
            if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Single Item: User choice set to $Choice" }
        }
        # If menu entry >=2
        else {
            $Choice = $($List[$Selection - 1])
            if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Multiple Item: User choice set to $Choice" }
        }
    }
    # If Selection is alpha
    else {
        $Choice = $Selection
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Alpha: User choice set to $Choice" }
    }

    return $Choice;
}

function Construct_Iventory() {
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Entered Construct_Inventory" }
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Contents of menu-entries.csv = $List" }

    <#-------------------------
        Create Root Directory
    --------------------------#>
    if (!(Test-Path "$PSScriptRoot\$InvDir\")) {
        New-Item -Path . -ItemType Directory -Name "$InvDir" | Out-Null
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Created $InvDir directory" }
    }
      
    <#-------------------------
        Create Menu
    --------------------------#>
    foreach ($Classroom in $List) {
        <#-------------------------
            Initialize Helpers
        --------------------------#>
        $TEMP_DIR = $Classroom.ToLower().replace(" ", "_")
        $TEMP_FILE = $Classroom.ToLower().replace(" ", "-") + '-computers' + '.txt'
        $FULL_PATH = "$PSScriptRoot\$InvDir\$TEMP_DIR\$TEMP_FILE"
        
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: TEMP_DIR initialized to $TEMP_DIR" }
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: TEMP_FILE initialized to $TEMP_FILE" }
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: FULL_PATH initialized to $FULL_PATH" }
          
        <#-------------------------
            Create Root\Directories
        --------------------------#>
        New-Item -Path "$PSScriptRoot\$InvDir\" -ItemType Directory -Name $TEMP_DIR -Force | Out-Null
  
        <#-------------------------
            Generate computer.txt files
        --------------------------#>
        if (!(Test-Path $FULL_PATH)) {
            $TEMP_RETVAL = $true  
            New-Item -Path "$PSScriptRoot\$InvDir\$TEMP_DIR\" -ItemType File -Name $TEMP_FILE | Out-Null
            Write-Host "Please update $TEMP_FILE in $InvDir\$TEMP_DIR with your computer hostnames.`n" -ForegroundColor Yellow;
            if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Script halted, new menu entry added" }
        }

        <#-------------------------
            Generate non-discoverable files
        --------------------------#>
        foreach ($File in $NonDiscoverablesFile) {
            if (!(Test-Path "$PSScriptRoot\$InvDir\$TEMP_DIR\$File")) {
                # Create the file
                New-Item -Path "$PSScriptRoot\$InvDir\$TEMP_DIR\" -ItemType File -Name $File | Out-Null
                if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: New file created @ $PSScriptRoot\$InvDir\$TEMP_DIR\$File" }

                # Add specific headers to the file
                if ($File -eq "projectors.csv") {
                    Set-Content "$PSScriptRoot\$InvDir\$TEMP_DIR\$File" -Value "Asset,Brand,Model,Model#,SerialNumber,ResolutionType,Resolution,Lumens,BulbStatus,LampCode,Status,History,Age,CurrentAge,Notes"
                    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Headers added to projectors.csv" }
                }
                elseif ($File -eq "apple.csv") {
                    Set-Content "$PSScriptRoot\$InvDir\$TEMP_DIR\$File" -Value "Asset,Brand,Model,ModelNum,SerialNumber,MAC,Age,CurrentAge"
                    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Headers added to apple.csv" }
                }
                elseif ($File -eq "printers.csv") {
                    Set-Content "$PSScriptRoot\$InvDir\$TEMP_DIR\$File" -Value "Asset,Brand,Model,ModelNum,SerialNumber,IP,Protocol,Age,CurrentAge"
                    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Headers added to printers.csv" }
                }
                elseif ($File -eq "miscellaneous.csv") {
                    Set-Content "$PSScriptRoot\$InvDir\$TEMP_DIR\$File" -Value "Asset,Category,Brand,Model,ModelNum,SerialNumber,Age,CurrentAge"
                    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Headers added to misc" }
                }
                else {
                    Set-Content "$PSScriptRoot\$InvDir\$TEMP_DIR\$File" -Value "Date,History"
                    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Default headers added to other files" }
                }

                # If anything in this loop was executed; prep exit message
                $TEMP_RETVAL = $true
                Write-Host "$File has been added to $InvDir\$TEMP_DIR, please populate it." -ForegroundColor Yellow; continue
                if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: RETVAL set to $TEMP_RETVAL; exiting script" }
            }
        }

        <#-------------------------
            Validate Content
        --------------------------#>
        if ($Null -eq (Get-Content $FULL_PATH)) {
            $TEMP_RETVAL = $true
            Write-Host "Please add content to $TEMP_FILE when you get the chance." -ForegroundColor Yellow; continue
            if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: No content found for $TEMP_FILE; exiting script" }
        }
    }
    <#-------------------------
        If NULL data; exit
    --------------------------#>
    if ($TEMP_RETVAL) {
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: TEMP_RETVAL set to $TEMP_RETVAL; exiting" }
        exit
    }
}

if ($Help) { Help; exit }

if (Test-Path "$PSScriptRoot\debug.log") {
    Remove-Item "$PSScriptRoot\debug.log" 
}

if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: Start of script" }

<#-------------------------
        Begin
--------------------------#>
$NonDiscoverablesFile = @("printers.csv", "projectors.csv", "apple.csv", "miscellaneous.csv", "history.csv", "todo.csv")
$InvDir = "Classrooms"
$InvFiles = Get-ChildItem "$PSScriptRoot\$InvDir\*"
$HTMLHeader = Get-Content "$PSScriptRoot\assets\style2.html"

# Array storage for Properties hashtable
$Results = @()

$ErrorActionPreference = 'Stop'

<#-------------------------
    Grab all menu entries
--------------------------#>
if (!(Test-Path "$PSScriptRoot\menu-entries.csv")) { 
    New-Item -Path "$PSScriptRoot\" -ItemType File -Name "menu-entries.csv" | Out-Null
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: No menu entries file exist; exiting script" }
    Write-Host "You had no menu entries file, please populate menu-entries.csv now" -ForegroundColor Yellow; exit
}
elseif ($Null -eq (Get-Content "$PSScriptRoot\menu-entries.csv")) {
    if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: You are missing content for menu-entries.csv; exiting script" }
    Write-Host "You are missing content for menu-entries.csv, please populate it now." -ForegroundColor Yellow; exit
}
else { $List = Get-Content "$PSScriptRoot\menu-entries.csv" | Sort-Object }

<#-------------------------
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
            $Shop = $Choice.ToLower().replace(" ", "_")
            $ShopPC = "$Choice-computers.txt".ToLower().Replace(" ", "-")

            # Match if $list array contains the $selected $Choice object
            Write-Host "`nYou chose $($Choice)." -ForegroundColor Green
            Write-Host "Inventory will be pulled from $Shop\$ShopPC`n" -ForegroundColor Green
        }
        'a' { 
            $File = "$PSScriptRoot\menu-entries.csv"

            Remove_WhiteSpace($File)

            # Ask for a selection; then sanitize it
            $Addition = Read-Host "What would you like to add?"
            $Addition = (Get-Culture).TextInfo.ToTitleCase($Addition)

            # Add the sanitized selection; prefix new line
            Add-Content -Path $File -Value `n$Addition

            Remove_WhiteSpace($File)

            exit
        }
        'b' {
            $TempKey = Read-Host "What computer would you like to scan?"
            Local_Scan($TempKey.ToUpper())
            exit
        }
        'n' {
            Write-Host "Non-Discoverable Devices" -ForegroundColor Yellow
            $Entry = 1
            foreach ($Item in $NonDiscoverablesFile) {
                Write-Host "[$Entry]: $($Item)"
                $Entry++
            }

            $Option = Read-Host "Choose a file to run a report on"
            $Option = $NonDiscoverablesFile[$Option-1]

            NonDiscoverable_Scan($Option)
            exit
        }
        'q' { exit }
        '\' { Write-Host "dev" }
        default { Write-Host "Invalid menu choice" -ForegroundColor Red }  
    }
} while ($Choice -eq 'q')

<#-------------------------
    Start Execution
--------------------------#>
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

<#-------------------------
        Computers
--------------------------#>
$Computers = (Get-Content "$PSScriptRoot\$InvDir\$Shop\$ShopPC")

<#-------------------------
    Directory/File Creation
--------------------------#>
$Filename = Read-Host "What would you like your filename to be? "
$Dirname = New-Item -Path "$PSScriptRoot\$InvDir\$Shop\" -Name "inv" -ItemType Directory -Force

$HTMLOut = "$Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html"
$CSVOut = "$Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv"

foreach ($Computer in $Computers) {
    try {
        <#-------------------------
            Properties Object/Hash
        --------------------------#>
        $Properties = @{
            Asset = ""
            Hostname = (Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).CSName
            Manufacturer = (Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem).Manufacturer
            Model = (Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem).Model
            Serial = (Get-WmiObject -ComputerName $Computer -Class Win32_Bios).SerialNumber
            Edition = (Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Caption
            OS = (Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem).Version
            Memory = Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem | Select-Object TotalPhysicalMemory, @{name="GB";expr={[float]($_.TotalPhysicalMemory/1GB)}} | Select-Object -ExpandProperty GB
            IP = ([System.Net.Dns]::GetHostByName("$Computer").AddressList[0]).IpAddressToString
            Domain = (Get-WmiObject -ComputerName $Computer -Class Win32_Computersystem).Domain
            MAC = (Get-WmiObject -ComputerName $Computer -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'").MACAddress | Select-Object -First 1
            Age = Get-WmiObject -ComputerName $Computer -Class Win32_BIOS
            ReimageDate = Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem
        }

        <#-------------------------
            ReleaseID Conversion
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
    catch {
        Write-Host "`n$Computer Offline" -ForegroundColor Red
        if ($Development) { Write-Log  "[$(GCLN)][DEBUG]: cannot reach $Computer" }
        continue
    }
}

<#-------------------------
        CSV Export
--------------------------#>
if ($Simple) {
    $Results | Select-Object Asset,Hostname,Manufacturer,Model,Serial,Edition,OS | Export-Csv -Path $CSVOut -NoTypeInformation
}
else {
    $Results | Select-Object Asset,Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | Export-Csv -Path $CSVOut -NoTypeInformation
}

<#-------------------------
        HTML Export
--------------------------#>
$PCTitle = "<div class='Header'>Computer Information Report for $Key</div>"

# Create the HTML file
if ($Simple) {
    $Results | Select-Object Asset,Hostname,Manufacturer,Model,Serial,Edition,OS | ConvertTo-Html -Head $HTMLHeader -Body $PCTitle | Out-File $HTMLOut
}
else {
    $Results | Select-Object Asset,Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | Sort-Object Hostname | ConvertTo-Html -Head $HTMLHeader -Body $PCTitle | Out-File $HTMLOut
}

foreach ($Item in $NonDiscoverablesFile) {
    $Csv = Import-Csv "$PSScriptRoot\$InvDir\$Shop\$Item"
    $History = $Item 
   
    # Strip the extension; Capitalize the first letter
    $Item = [IO.Path]::GetFileNameWithoutExtension($Item); $Item = (Get-Culture).TextInfo.ToTitleCase("$Item".ToLower())

    if ($Csv.length -eq 0) { continue }
    elseif ($History -eq 'history.csv') {
        "<br><br><div class='Header'>$Item</div>" | Out-File $HTMLOut -Append
        $Csv | Select-Object * | Sort-Object -Property Date | ConvertTo-Html -Fragment | Out-File $HTMLOut -Append
    }
    else {
        "<br><br><div class='Header'>Information Report for $Item</div>" | Out-File $HTMLOut -Append
        $Csv | Select-Object * | Sort-Object -Property Asset | ConvertTo-Html -Fragment | Out-File $HTMLOut -Append
    }
}

# Append footer
"<div class='Copyright'>$((Get-Date -Format g))</div>
</html>" | Out-File $HTMLOut -Append

<#-------------------------
    Report Locations
--------------------------#>
Write-Host "You can view the logfile at $CSVOut" -ForegroundColor Green
Write-Host "You can view the HTML report at $HTMLOut" -ForegroundColor Green

<#-------------------------
    Script Execution
--------------------------#>
$Stopwatch.Stop()
$Time = $Stopwatch.Elapsed
Write-Host "`nThe script completed in $Time seconds`n"