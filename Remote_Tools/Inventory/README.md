# Litter-Box

## Startup

1. Clone this repo
2. Navigate to repo
3. Execute .\Litter-Box.ps1
4. Follow script instructions

## The Specifics

This script is designed to take inventory of computers in your domain.

It starts by dynamically generating a file called `menu-entries.csv`. From here, it will notify the user to update that file with a *Classroom* (i.e. the directory that holds whatever computers you want to inventory), (I just work for a school so it happens to be Classrooms), (The script isn't hard to modify to make it say something else)

After menu-entries is populated, the script must be executed again. This time it will warn the user that they need to populate a text file called [menu-entry]-computers.txt. Place your computer hostnames in that file.

Execute the script for a third time (I promise this is the last time) and you will now see a menu generated based on your `menu-entries.csv` file. Choose the appropriate option. The script will notify you of where you can find the inventory reports, which are both in `.csv` and `.html`.

The startup can be a little annoying, but once you have at least one menu item, adding to it is pretty self explanatory. The warnings are only in place so that you aren't generating inventory of nothing or causing the script to make an oopsie-woopsie.

## Future Improvements

Adding functionality for "non-discoverable" devices is next on the list. Obviously, WMI cannot scan everything. Some items I need to inventory are projectors, Apple devices, A/V equipment, etc. I dabbled with this in version 1, but I was more worried about the dynamic menu at that point.

Hopefully at some time in the future we will see a better looking HTML table. It tends to be a little tricky to generate a nice looking table with the appropriate CSS.

It would be nice to see some type of AD integration, as well. Something that can automatically generate a `computers.txt` file for the user. However, this just doesn't fit my use-case at this time and honestly probably never will. Feel free to modify the script yourself.

I'm no PowerShell guru, but I've read some stuff about WMI not be the "best" thing since toast, so maybe there is a better implementation that someone can generate. I feel that WMI was the easiest to use since WinRM/PSRemoting is not needed. Other methods I've seen used CIM instances.

## Regards

I personally use this script almost daily, whether it be to see what version Windows is running on a specific computer, showing administrators what classroom technology looks like, checking the age of the computers, etc.

If anyone has questions, feel free to message me somehow or create issues/enhancements.