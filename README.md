# Common VBA File Services
## Services
### Common

| Name           | Service                                    |
| -------------- | ------------------------------------------ |
| Arry           | Get: Returns the content of a text file as an array.|
|                | Let: Write the content of an array to a file |    
| Compare        | Function: Displays the differences between two files by means of WinMerge |
| Delete         | Sub: Deletes a file provided either as object or as full name when it exists  |
| Dict           | Function: Returns the content of a test file as Dictionary |
| Differs        | Function: Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option |
| Exists         | Function: Returns True when a given file exists, plus a collection of all files found when specified with wildcards |
| Extension      | Function: Returns a the extension of file's name. The file may be provided either as file object or as full name|
| GetFile        | Function: Returns a file object provided by its full name |
| Search         | Function: Returns a collection of all files found supporting wildcards and subfolders | 
| Sections.      | Returns the sections if a _Private Properties File_ as Dictionary with the _Section_ as the key and the _Values_ (see below) as the item |
| SelectFile     | Function: Returns the full name of a file selected in a dialog |
| Temp           | Property Get: Provides the full name of an arbitrary named file, by default in the current directory or in a given path with and optional extension which defaults to .tmp | 
| Txt            | Property Get: Provides the content of a text file as string, optionally with the split string for the VBA.Split operation which may be used to transfer the string into an array |
|                | Property Let: Writes a text string, optionally intermitted by vbCrLf, to a file - optionally appended. |

### PrivateProfile services
Simplify the handling of .ini, .cfg, or any other file organized by [section] value-name=value. Consequently all services primarily have the arguments. file, section, and value-name. In order to extend the possible usages of such a file some extra services are provided. 

| Name           | Service                                      |
| -------------- | -------------------------------------------- |
| NameRemove     | Sub: Removes a named value entry from a file |
|                | Syntax:<br>`mFile.NameRemove file, section, valuename` |
| SectionExists  | Returns True when a given section [.....] exists in the file. |
|                | Syntax:<br>`If mFile.SectionExists(file, section) Then ...`|
| SectionsCopy   | Sub: Uses Sections Get/Let to copy named - or when omitted all - sections from a source to a target file, with the target file sections optionally replaced, by default is merged. When all sections are copied (i.e. no section names are provided), the option replace is used, and the target file is identical with the source file the sections will only be reorganized in ascending order. |
|                | Syntax:<br>`mFile.SectionsCopy source-file, target-file[, sections][, replace]`|
| SectionsRemove | Sub: Removes the named or all sections [.....] from a file |
|                | Syntax:<br>`mFile.SectionsRemove file, section`
| Value          | Get: Returns the value for a given name from a given section from a _Private Properties File_ |
|                | Let: write a value with a given name under a given section in a _Private Properties File_ |
| Values         | Returns the values from a given _Private Properties File_ under a given _Section_ |
   

## Installation
Download and import [mFile.bas][1] to your VB project.
Download and import [mDct.bas][2] to your VB project.

## Usage
See table above

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/source/mFile.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Directory-Services/master/source/mDct.bas
