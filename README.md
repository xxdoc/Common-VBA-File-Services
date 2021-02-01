# Common VBA File Services
## Services
### Common

| Name           | Service                                    |
| -------------- | ------------------------------------------ |
| Arry           | Property Get: the content of a text file as an array.|
| Compare        | Function: Displays the differences between two files by means of WinMerge |
| Dct            | Property Get: Returns the content of a text file as Dictionary |
| Delete         | Sub: Deletes a file when existing with the file provided either as object or as full name |
| Differs        | Function: Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option |
| Exists         | Function: Returns True when a given file exists, plus a collection of all files found when specified by wildcards |
| Extension      | Function: Returns the extension of a file provided either as file object or full name
| GetFile        | Function: Returns a file object provided by its full name |
| Search         | Function: Returns a collection of full file names | 
| SelectFile     | Function: Returns the full name of a file selected in a displayed dialog |
| Temp           | Property Get: Provides the full name of an arbitrary named file in the users Temp directory, with an optional extension which defaults to .tmp | 
| Txt            | Property Get: Provides the content of a text file as string, optionally with the split string for the VBA.Split operation which may be used to transfer the string into an array<br>Property Let: Writes a text string, optionally intermitted by vbCrLf, to a file - optionally appended. |

### PrivateProfile services

| Name           | Service                                      |
| -------------- | -------------------------------------------- |
| NameRemove     | Sub: Removes a named value entry from a file |
| SectionNames   | Property Get: Provides all section [.....] names in a file in as ending order as Collection. |
| Sections    | Property Get: Returns the named or all sections [......] in a file in a Dictionary with the section name as the key and the item as a Dictionary of all names and values<br>Property Let: Writes the sections provided as Dictionary to a file|
| SectionsCopy   | Sub: Uses Sections Get/Let to copy named - or when omitted all - sections from a source to a target file, with the target file sections optionally replaced, by default is merged. When all sections are copied (i.e. no section names are provided), the option replace is used, and the target file is identical with the source file the sections will only be reorganized in ascending order. |
| SectionsRemove | Sub: Removes the named or all sections [.....] from a file |
| Value          | Property Get: the named value in a file with sections and name=value records<br>Property Let: write a named value to a file |
| ValueNames     | Function: Returns all value names, when no section name is provided of all [.....] sections in a file, in ascending order with no duplicates as Collection (uses the module DctAdd service of the mDct module) |
| Values         | Function: Provides all values of the named or all [....] sections as Dictionary in ascending order with duplicates ignored, the value name as key and the value as item (uses the DctAdd service in module mDct)|

## Installation
Download and import [mFile][1] to your VB project.

## Usage
See table above

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles are appreciated.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas