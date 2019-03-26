# Batch-File-Rename

## Description

This script can rename files to a specific name located in a CSV file.

## Alternative Use

Before reading my use-case for this script, here are a few things you can do with it.

Lines 35-44 are the tuna and potatoes of this script. It reads in the `photo-list.csv` file using `Import-Csv`. This reads as a type of `Array`. What this means is a simple for loop can turn a specific process into something simple.

Because CSV files are treated as arrays, you can use different methods to access different data. If you have a different CSV file with different fields, it is as simple as changing `line 36` to your column name and `line 8` to import your different CSV name.

`Rename-Item` can be replaced with whatever command you desire. It just so happens that this cmdlet takes in two parameters. If you're curious to see a similar implementation, check out my [Rename Computer](https://github.com/importedtea/powershell/tree/master/Remote_Tools/RenameComputer) directory

As for other parts of the script that backup the images... Well, you can remove those parts if you want. They are a pretty basic/poor implementation of backing up the directories. It obviously won't work for every scenario so you might want to modify those parts. But honestly, just remove them because they are pretty specific to my situation.

## Purpose

I wrote this script to sanitize image file names for a Student Information System called ClassMate. The SIS only accepts images with a specific file name that just so happens to be the student ID number. However, the camera that the images are taken on increments through filenames like P17091, P17092, etc.

You end up having a directory of images like this:

| Filename |
|---|
|P17091|
|P17092|
|P17093|

## Exporting Information from ClassMate

The best method to get all the student data is through an export of the system. You can specify a few fields, mainly ID, First, Last, and Course. The only field needed is ID in this scenario, but for those taking the pictures, those other fields help them. It looks like the below:

| ID | First | Last | Course |
|---|---|---|---|
| 000123123 | Student  | A | Course#1  |
| 000123124 | Student  | B | Course#1  |
| 000123125 | Student  | C | Course#1  |

## The Process

This whole image process goes in three phases. Phase 1 needs to be very particular and must go in order. If it strays from the list, it must be noted so that IT can remove this student from the master list before running the script.

**Phase 1:** Staff member is given the student export file, and takes photos in order of student name.

**Phase 2:** The camera is given to IT and the images are removed and placed in `IMAGES\`

**Phase 3:** The script is then executed. The directory structure must contain all files in this repo to work.