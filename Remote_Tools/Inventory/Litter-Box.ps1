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

Future Implementation:
Get a list of computers from AD (priority = -1)
~~Add an alpha option to Add-Content to menu-entries.csv (maybe do this method instead, or both)~~
If an age field is given in a csv, use that to calculate based on current time/date. (ex. printers.csv has Age column that says 0 (i.e. new), calculate against Get-Date to verify age when script was executed)
~~Add compatibility for a company logo at the top of the screen.~~
~~Add a help function and paramter -Help~~
Add functionality to grab only specific devices. For instance, grab all projector data from Classrooms\*; Add as alpha option
Generate non-discoverable files when a menu-entry is added, elimminating one script execution
Add a variable for Classrooms for easy change of directory name; plus the word is used in a lot of paths of the script
Condense static variables into one location for ease of use

.LINK
    https://github.com/importedtea/powershell/tree/master/Remote_Tools/Inventory
#>

param (
    [Parameter(Mandatory=$false)]
    [switch] $Simple,
    [Parameter(Mandatory=$false)]
    [switch] $Help
)

function Help() {
    Write-Host @"
        
        Running the script:
            .\Litter-Box
        
        Running simplified version:
            .\Litter-Box -Simple

        Launching Help:
            .\Litter-Box -Help
        
        Menu-Entries.csv
            Step 1: menu-entries.csv is added on initial run after clone.
            Step 2: script will notify you that the file was added.
            Step 3: if you try running it again, it will tell you it is empty.
            Step 4: open menu-entries.csv and add a domain to store your inventory (lower | uppercase)

        After Content is Added:
            Step 1: when you run the script, it will notify you that you added something 
                and prompt you to add hostnames.
            Step 2: when you run it again, it will generate the non-discoverable files.
            Step 3: Add a hostname to *-computers.txt file.
            Optional: Add non-discoverable information in appropriate .csv file.
            Step 4: .\Litter-Box.ps1

        Choosing Your Content:
            If you followed all the warning prompts, you should successfully see a menu
                with your item and two additional alpha values.
            Step 1: Choose your menu entry
            Step 2: Give your inventory output file a name
            Step 3: Watch the magic happen

        Adding Content:
            You have two choices of adding content. You can either use the alpha a|A option in the menu
                or manually update the menu-entries.csv file.
        
            Either option supports lower|uppercase. It is possible to break this so try to avoid special chars.

"@
}

function GCLN { $MyInvocation.ScriptLineNumber } 

function Write-Log($Message, $Path="$PSScriptRoot\debug.log") {
    "[$(Get-Date -Format g)] => $Message" | Tee-Object -FilePath $Path -Append | Write-Verbose
}

function Remove_WhiteSpace($File) {
    # Grab file contents; remove any trailing whitespace
    $Newtext = (Get-Content -Path $File -Raw) -replace "(?s)`r`n\s*$"
    [System.IO.File]::WriteAllText($File,$Newtext)
}

function Show-Menu {
    Write-Log "[$(GCLN)][DEBUG]: Entered Show-Menu function"
    # Clear-Host

    <#-------------------------
        Print Menu Title
    --------------------------#>
    Write-Host "`n************ MENU ************"
    Write-Log "[$(GCLN)][DEBUG]: Wrote Title"

    <#-------------------------
        Print Menu Entries
    --------------------------#>
    $Entry = 1
    Write-Log "[$(GCLN)][DEBUG]: Initialized menu numbers"
    foreach ($Item in $List) {
        Write-Host "[$Entry]: $($Item)"
        $Entry++
    }
  
    <#-------------------------
        Add Static Entries
    --------------------------#>
    Write-Host "`n[A]: Add Menu Entry"
    Write-Host "[Q]: Quit`n"
    
    Write-Log "[$(GCLN)][DEBUG]: Appended add option"
    Write-Log "[$(GCLN)][DEBUG]: Appended quit option"
      
    <#-------------------------
        Accept User Input
    --------------------------#>
    $Selection = Read-Host "Select a menu option"
    Write-Log "[$(GCLN)][DEBUG]: Prompted user for input; user chose $Selection"
  
    <#-------------------------
        Validate STDIN
    --------------------------#>
    # If Selection is numeric
    if ($Selection -match "^\d+$") {
        # If one menu item, read that item
        if ((Get-Content "$PSScriptRoot\menu-entries.csv").Count -eq 1) {
            $Choice = $List
            Write-Log "[$(GCLN)][DEBUG]: Single Item: User choice set to $Choice"
        }
        # If menu entry >=2
        else {
            $Choice = $($List[$Selection - 1])
            Write-Log "[$(GCLN)][DEBUG]: Multiple Item: User choice set to $Choice"
        }
    }
    # If Selection is alpha
    else {
        $Choice = $Selection
        Write-Log "[$(GCLN)][DEBUG]: Alpha: User choice set to $Choice"
    }

    <#-------------------------
        Return the Choice
    --------------------------#>
    return $Choice;
}

function Construct_Iventory() {
    Write-Log "[$(GCLN)][DEBUG]: Entered Construct_Inventory"
    Write-Log "[$(GCLN)][DEBUG]: Contents of menu-entries.csv = $List"
    <#-------------------------
        Create Root Directory
    --------------------------#>
    if (!(Test-Path "$PSScriptRoot\Classrooms\")) {
        New-Item -Path . -ItemType Directory -Name "Classrooms" | Out-Null
        Write-Log "[$(GCLN)][DEBUG]: Created Classrooms directory"
    }
      
    <#-------------------------
        Create Menu Loop
    --------------------------#>
    foreach ($Classroom in $List) {
        <#-------------------------
            Initialize Helpers
        --------------------------#>
        $TEMP_DIR = $Classroom.ToLower().replace(" ", "_")
        $TEMP_FILE = $Classroom.ToLower().replace(" ", "-") + '-computers' + '.txt'
        $FULL_PATH = "$PSScriptRoot\Classrooms\$TEMP_DIR\$TEMP_FILE"
        
        Write-Log "[$(GCLN)][DEBUG]: TEMP_DIR initialized to $TEMP_DIR"
        Write-Log "[$(GCLN)][DEBUG]: TEMP_FILE initialized to $TEMP_FILE"
        Write-Log "[$(GCLN)][DEBUG]: FULL_PATH initialized to $FULL_PATH"
          
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
            Write-Log "[$(GCLN)][DEBUG]: Script halted, new menu entry added"
        }

        <#-------------------------
            Generate non-discoverable files
        --------------------------#>
        $global:NonDiscoverablesFile = @("printers.csv", "projectors.csv", "apple.csv", "miscellaneous.csv")

        foreach ($File in $NonDiscoverablesFile) {
            if (!(Test-Path "$PSScriptRoot\Classrooms\$TEMP_DIR\$File")) {
                # Create the file
                New-Item -Path "$PSScriptRoot\Classrooms\$TEMP_DIR\" -ItemType File -Name $File | Out-Null
                Write-Log "[$(GCLN)][DEBUG]: New file created @ $PSScriptRoot\Classrooms\$TEMP_DIR\$File"

                # Add specific headers to the file
                if ($File -eq "projectors.csv") {
                    Set-Content "$PSScriptRoot\Classrooms\$TEMP_DIR\$File" -Value "Asset,Brand,Model,Model#,SerialNumber,ResolutionType,Resolution,Lumens,BulbStatus,LampCode,Status,History,Age,CurrentAge"
                    Write-Log "[$(GCLN)][DEBUG]: Headers added to projectors.csv"
                }
                elseif ($File -eq "apple.csv") {
                    Set-Content "$PSScriptRoot\Classrooms\$TEMP_DIR\$File" -Value "Asset,Brand,Model,ModelNum,SerialNumber,MAC,Age,CurrentAge"
                    Write-Log "[$(GCLN)][DEBUG]: Headers added to apple.csv"
                }
                elseif ($File -eq "printers.csv") {
                    Set-Content "$PSScriptRoot\Classrooms\$TEMP_DIR\$File" -Value "Asset,Brand,Model,ModelNum,SerialNumber,IP,Protocol,Age,CurrentAge"
                    Write-Log "[$(GCLN)][DEBUG]: Headers added to printers.csv"
                }
                else {
                    Set-Content "$PSScriptRoot\Classrooms\$TEMP_DIR\$File" -Value "Asset,Category,Brand,Model,ModelNum,SerialNumber,Age,CurrentAge"
                    Write-Log "[$(GCLN)][DEBUG]: Default headers added to other files"
                }

                # If anything in this loop was executed; prep exit message
                $TEMP_RETVAL = $true
                Write-Host "You have added $File, please populate it." -ForegroundColor Yellow; continue
                Write-Log "[$(GCLN)][DEBUG]: RETVAL set to $TEMP_RETVAL; exiting script"
            }
        }

        <#-------------------------
            Validate Content
        --------------------------#>
        if ($Null -eq (Get-Content $FULL_PATH)) {
            $TEMP_RETVAL = $true
            Write-Host "You are missing content for $TEMP_FILE, please populate it now." -ForegroundColor Yellow; continue
            Write-Log "[$(GCLN)][DEBUG]: No content found for $TEMP_FILE; exiting script"
        }
    }
    <#-------------------------
        If NULL data; exit
    --------------------------#>
    if ($TEMP_RETVAL) { Write-Log "[$(GCLN)][DEBUG]: TEMP_RETVAL set to $TEMP_RETVAL; exiting"; exit }
}

if ($Help) { Help; exit }

Remove-Item "$PSScriptRoot\debug.log" 
Write-Log "[$(GCLN)][DEBUG]: Start of script"

<#-------------------------
    Grab all menu entries
--------------------------#>
if (!(Test-Path "$PSScriptRoot\menu-entries.csv")) { 
    New-Item -Path "$PSScriptRoot\" -ItemType File -Name "menu-entries.csv" | Out-Null
    Write-Log "[$(GCLN)][DEBUG]: No menu entries file exist; exiting script"
    Write-Host "You had no menu entries file, please populate menu-entries.csv now" -ForegroundColor Yellow; exit
}
elseif ($Null -eq (Get-Content "$PSScriptRoot\menu-entries.csv")) {
    Write-Log "[$(GCLN)][DEBUG]: You are missing content for menu-entries.csv; exiting script"
    Write-Host "You are missing content for menu-entries.csv, please populate it now." -ForegroundColor Yellow; exit
}
else { $List = Get-Content "$PSScriptRoot\menu-entries.csv" }
  
<#-------------------------
    Grab HTML Styling
--------------------------#>
$Header = Get-Content "$PSScriptRoot\assets\style2.html"

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

            # Remove any whitespace
            Remove_WhiteSpace($File)

            # Ask for a selection; then sanitize it
            $Addition = Read-Host "What would you like to add?"
            $Addition = (Get-Culture).TextInfo.ToTitleCase($Addition)

            # Add the sanitized selection; prefix new line
            Add-Content -Path $File -Value `n$Addition

            # Remove any whitespace
            Remove_WhiteSpace($File)

            # exit needs to occur
            exit
        }
        'q' { exit }
        '\' { Write-Host "dev" }
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
            Properties Object/Hash
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
    <#-------------------------
        Error Handling
    --------------------------#>
    catch {
        Write-Host "`n$computer Offline" -ForegroundColor Red
        Write-Log "[$(GCLN)][DEBUG]: cannot reach $computer"
        continue
    }
}

# The below calculates the age field against todays date, however, it ends up breaking any futher script execution; just a note

# ### Calculate CSV age
# # Get the csv file names
# $Files = Get-ChildItem "$PSScriptRoot\Classrooms\$Shop\*.csv"

# # Loop through the filenames
# foreach ($Filename in $Files) {
#     $Filename
    
#     $CSV = Import-Csv $Filename

#     # Set counter to 0; reset with each new csv file
#     $Line = 0

#     # Loop through each csv data row
#     foreach ($CC in $CSV.Age) {
#         Write-Host "Line = $Line"
#         Write-Host "Line $Line : CSV = $CSV"
#         Write-Host "Made it to nest: $CC"
#         # Get todays date
#         $Today = Get-Date -UFormat "%m/%d/%Y"
#         Write-Host "Todays date = $Today"

#         # Grab the first file; import csv into $CSV
#         $V = [datetime]$CSV[$Line].Age
#         Write-Host "CSV age = $V"
        
#         $Diff = New-TimeSpan -Start $V -End $Today
#         $Diff.Days / 365
#         Write-Host `n`n
#         $Line++
#     }
# }

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
$PCTitle = "<div class='header'>Computer Information Report for $Key</div>"

# Create the HTML file
if ($Simple) {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS | ConvertTo-Html -Head $Header -Body $PCTitle | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html
}
else {
    $Results | Select-Object Hostname,Manufacturer,Model,Serial,Edition,OS,Memory,IP,MAC,Domain,Age,ReimageDate | Sort-Object Hostname | ConvertTo-Html -Head $Header -Body $PCTitle | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html
}

foreach ($Item in $NonDiscoverablesFile) {
    $Csv = Import-Csv "$PSScriptRoot\Classrooms\$Shop\$Item"
    
    # Strip the extension; Capitalize the first letter; (No && in PS)
    $Item = [IO.Path]::GetFileNameWithoutExtension($Item); $Item = (Get-Culture).TextInfo.ToTitleCase("$Item".ToLower())

    if ($Csv.length -eq 0) { continue }
    else {
        "<br><br><div class='header'>Information Report for $Item</div>" | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html -Append
        $Csv | Select-Object * | Sort-Object -Property Asset | ConvertTo-Html -Fragment | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html -Append
    }
}

# Append a footer
"<div class='Copyright'>$((Get-Date -Format g))</div>
</html>" | Out-File $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html -Append

<#-------------------------
    Report Locations
--------------------------#>
Write-Host "You can view the logfile at $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).csv" -ForegroundColor Green
Write-Host "You can view the HTML report at $Dirname\${Filename}_$((Get-Date).ToString('MM-dd-yyyy')).html" -ForegroundColor Green

<#-------------------------
    Script Execution
--------------------------#>
$Stopwatch.Stop()
$Time = $Stopwatch.Elapsed
Write-Host "`nThe script completed in $Time seconds`n"