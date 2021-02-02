# Common VBA File Services
## Services
### Common

| Name           | Service                                    |
| -------------- | ------------------------------------------ |
| Arry           | Property Get: the content of a text file as an array.|
| Compare        | Function: Displays the differences between two files by means of WinMerge |
| Delete         | Sub: Deletes a file provided either as object or as full name when it exists  |
| Differs        | Function: Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option |
| Exists         | Function: Returns True when a given file exists, plus a collection of all files found when specified with wildcards |
| Extension      | Function: Returns a the extension of file's name. The file may be provided either as file object or as full name|
| GetFile        | Function: Returns a file object provided by its full name |
| Search         | Function: Returns a collection of all files found supporting wildcards and subfolders | 
| SelectFile     | Function: Returns the full name of a file selected in a dialog |
| Temp           | Property Get: Provides the full name of an arbitrary named file, by default in the current directory or in a given path with and optional extension which defaults to .tmp | 
| Txt            | Property Get: Provides the content of a text file as string, optionally with the split string for the VBA.Split operation which may be used to transfer the string into an array |
|                | Property Let: Writes a text string, optionally intermitted by vbCrLf, to a file - optionally appended. |

### PrivateProfile services
Simplify the handling of .ini, .cfg, or any other file organized by [section] value-name=value. Consequently all services primarily have the arguments. file, section, and value-name. In order to extend the possible usages of such a file some extra services are provided. 

| Name           | Service                                      |
| -------------- | -------------------------------------------- |
| NameRemove     | Sub: Removes a named value entry from a file |
|                | Syntax:<br>`mFile.NameRemove file, section, value-name` |
| SectionNames   | Property Get: Provides all section [.....] names in a file in as ending order as Collection. |
|                | Syntax get the names:<br>`Set dct = mFile.SectionNames(file)`|
|                | Syntax check a section exists:<br>`If mFile.SectionNames(file).Exists(section-name) then ...`
| Sections       | Property Get: Returns the named or all sections [......] in a file in a Dictionary with the section name as the key and the item as a Dictionary of all names and values |
|                | Syntax:<br>`Set dct = mFile.Sections(file[, section-names]` |
|                | Property Let: Writes the sections provided as Dictionary to a file|
|                | Syntax:<br>`mFile.Sections(file) = dct` |
| SectionsCopy   | Sub: Uses Sections Get/Let to copy named - or when omitted all - sections from a source to a target file, with the target file sections optionally replaced, by default is merged. When all sections are copied (i.e. no section names are provided), the option replace is used, and the target file is identical with the source file the sections will only be reorganized in ascending order. |
|                | Syntax:<br>`mFile.SectionsCopy source-file, target-file[, sections][, replace]`|
| SectionsRemove | Sub: Removes the named or all sections [.....] from a file |
| Value          | Property Get: the named value in a file with sections and name=value records<br>Syntax:<br>`value = mFile.Value(file,section,value-name)` |
|                | Property Let: write a named value to a file<br>Syntax:<br>`mFile.Value(file,section,value-name) = value` |
| ValueNames     | Function: Returns all value names, when no section name is provided of all [.....] sections in a file, in ascending order with no duplicates as Collection (uses the module DctAdd service of the mDct module)<br>Syntax get all value names in all or the provided sections:<br>`Set dct = mFile.ValueNames(file[, sections]`<br>Syntax check if a name exists in a provided or any section:<br>`If mFile.ValueNames(file[, sections].Exist(value-name)  Then ...`
| Values         | Function: Provides all values of the named or all [....] sections as Dictionary in ascending order with duplicates ignored, the value name as key and the value as item (uses the DctAdd service in module mDct)|
|                | Syntax:<br>`Set dct = mFile.Values(file, section, value-name)`<br>

## Installation
Download and import [mFile][1] to your VB project.

## Usage
See table above

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles are appreciated.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas