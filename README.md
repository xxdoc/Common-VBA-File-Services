# Common VBA File Services
## Services

| Name           | Service                                    | Named arguments | Meaning |
| -------------- | ------------------------------------------ | --------------- | ------- |
| Arry           | Property Get: the content of a text file as an array.|                 |         |
| Compare        | Function: Displays the differences between two files by means of WinMerge | | |
| Dct            | Property Get: Returns the content of a text file as Dictionary | | |
| Delete         | Sub: Deletes a file when existing with the file provided either as object or as full name | | |
| Differs        | Function: Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option | | |
| Exists         | Function: Returns True when a given file exists, plus a collection of all files found when specified by wildcards | | |
| Extension      | Function: Returns the extension of a file provided either as file object or full name
| GetFile        | Function: Returns a file object provided by its full name | | |
| NameRemove     | Sub: Removes a named value entry from a file | | |
| Search         | Function: Returns a collection of full file names | | |
| SectionNames   | Property Get: all [.....] section names in a file as Dictionary. | | |
| SectionsCopy   | Sub: Copies sections provided by their name from on file to another optionally merged or replaced | | |
| SectionsGet    | Function: Returns the sections [......] in file specified by their name in a Dictionary with the section name as the key and the item as a Dictionary of all names and values | | |
| SectionsLet    | Sub: Writes the sections provided as Dictionary to a file | | |
| SectionsRemove | Sub: Removes all named sections [.....] from a file | | |
| SelectFile     | Function: Returns a file selected in a displayed dialog | | |
| Temp           | Property Get: Provides the name of a temporary file, optionally with a certain extension | | | 
| Txt            | Property Get: the content of a text file as string<br>Property Let: Write a string to a file - optionally appended. | | |
| Value          | Property Get: the named value in a file with sections and name=value records<br>Property Let: write a named value to a file | | |
| ValueNames     | Function: Returns all names, when no section name is provided all of a file, in ascending sequence with duplicates ignored (requires the module mDct) | | |
| Values         | Function: Returns all values, when no section is provided all in a file, in a Dictionary in ascending order with duplicates ignored (requires the module mDct)

The named arguments and their meaning is pretty self explanatory but will be completed soon.

## Installation
Download and import [mFile][1] to your VB project.

## Usage
See table above

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles are appreciated.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas


