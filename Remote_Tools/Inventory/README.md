# Get-Inventory

## Startup

1. Clone this repo
2. Navigate to repo
3. Execute .\Get-Inventory.ps1
4. Follow script instructions

## The Specifics

This script is designed to take inventory of computers in your domain.

It starts by dynamically generating a file called `menu-entries.csv`. From here, it will notify the user to update that file with a *Classroom* (i.e. the directory that holds whatever computers you want to inventory).

After menu-entries is populated, the script must be executed again. This time it will warn the user that they need to populate a text file called [menu-entry]-computers.txt. Place your computer hostnames in that file.

Execute the script for a third time (I promise this is the last time) and you will now see a menu generated based on your `menu-entries.csv` file. Choose the appropriate option. The script will notify you of where you can find the inventory reports, which are both in `.csv` and `.html`.

The startup can be a little annoying, but once you have at least one menu item, adding to it is pretty self explanatory. The warnings are only in place so that you aren't generating inventory of nothing or causing the script to make an oopsie-woopsie.