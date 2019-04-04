# Litter-Box

## Startup

1. Clone this repo
2. Navigate to repo
3. Execute .\Litter-Box.ps1
4. Follow script instructions

## The Specifics

This script is designed to take inventory of computers in your domain.

It starts by dynamically generating a file called `menu-entries.csv`. From here, it will notify the user to update that file with a *Classroom* (i.e. the directory that holds whatever computers you want to inventory), (I just work for a school so it happens to be Classrooms), (The script isn't hard to modify to make it say something else)

After menu-entries is populated, the script must be executed again. This time it will warn the user that they need to populate a text file called [menu-entry]-computers.txt. Place your computer hostnames in that file. It will also generate a few non-discoverable .csv files.

Execute the script for a third time (I promise this is the last time) and you will now see a menu generated based on your `menu-entries.csv` file. Choose the appropriate option. The script will notify you of where you can find the inventory reports, which are both in `.csv` and `.html`.

The startup can be a little annoying, but once you have at least one menu item, adding to it is pretty self explanatory. The warnings are only in place so that you aren't generating inventory of nothing or causing the script to make an oopsie-woopsie. I recommend adding new menu items with the alpha a|A option once at least one menu entry exists.

Feel free to use `-Help` to get a short run down of some of the functionality.

## Future Improvements

~~Adding functionality for "non-discoverable" devices is next on the list. Obviously, WMI cannot scan everything. Some items I need to inventory are projectors, Apple devices, A/V equipment, etc. I dabbled with this in version 1, but I was more worried about the dynamic menu at that point.~~

I have an idea about changing the computer scanning process. The main idea behind it is to have a function to generate whatever computer scans are available at that moment. There would be another function that would compile the generated computers and add them to a "master" list. Maybe using something like XML or JSON would be cool. Basically, I want a non discoverable scan like function to be implemented for computers so that there isn't too much of a worry about overwriting something already there. When it comes down to it, you need permanent data and you can't always guarantee that computers will be pinged via WMI. You also don't want to have a list of 20 computers that were scanned, and then run it again and only have 5 on the list. Currently the script names files with a timestamp to distinguish any differences. You can probably see the point from all this rambling.

Hopefully at some time in the future we will see a better looking HTML table. It tends to be a little tricky to generate a nice looking table with the appropriate CSS.

I'm no PowerShell guru, but I've read some stuff about WMI not be the "best" thing since toast, so maybe there is a better implementation that someone can generate. I feel that WMI was the easiest to use since WinRM/PSRemoting is not needed. Other methods I've seen used CIM instances.

## Regards

My cat has a pretty organized inventory in her litter box. She's always brushing it around trying to find the one, and constantly adding new content daily. Truly fascinating. Plus it sounds better than some 46 year old naming convention like "Get-Inventory".

I started writing this script as something to do, and a way to boost my PowerShell skillset. Overall, it has been a lot of fun and I am constantly thinking of new implementations daily.

Probably one of my favorite features is the `history.csv` file. Anywhere you work there will be a lot of unknowns (*cough* especially in K-12 *cough*), it's sort of a pain to update the history file, but if you make a habit of it, you can generate a nice report in the end and see exactly when something was done.

If anyone has questions, feel free to message me somehow or create issues/enhancements.