## VBA File Services
### Common services

| Name           | Service                                    |
| -------------- | ------------------------------------------ |
| Arry           | Get: Returns the content of a text file as an array.|
|                | Let: Write the content of an array to a file |    
| Compare        | Function: Displays the differences between two files by means of WinMerge |
| Delete         | Deletes a file provided either as object or as full name when it exists  |
| Dict           | Returns the content of a test file as Dictionary |
| Differs        | Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option |
| Exists         | Returns True when a given folder, file, section, or value-name exists, returns a collection of all files found. Syntax:<br>`mFile.Exists(folder[, file[, section[, value-name]]]` |
| Extension      | Returns a the extension of file's name. The file may be provided either as file object or as full name|
| GetFile        | Returns a file object for given file's full name |
| Search         | Returns a collection of all files found supporting wildcards and sub-folders | 
| SelectFile     | Returns the full name of a file selected in a dialog |
| Temp           | Property Get: Provides the full name of an arbitrary named file, by default in the current directory or in a given path with and optional extension which defaults to .tmp | 
| Txt            | Get: Returns the content of a text file as string, returns the split string/character for the VBA.Split operation which may be used to transfer the string into an array |
|                | Let: Writes a text string, optionally intermitted by vbCrLf, to a file - optionally appended. |

### _PrivateProfile File_ services
Simplifies the handling of .ini, .cfg, or any other file organized by [section] value-name=value. Consequently all services use the following named arguments:

| Name           | Description                                                                     |
| -------------- | ------------------------------------------------------------------------------- |
| pp_file        | Variant expression, either a _PrivateProfile File's_ full name or a file object |
| pp_section     | String expression, the name of a _Section_                                      |
| pp_value_name  | String expression, the name of a _Value_ under a given _Section_                |
| pp_value       | Variant expression, will be written to the file as string                       |

| Name           | Service                                      |
| -------------- | -------------------------------------------- |
| NameRemove     | Sub: Removes a named value entry from a file |
|                | Syntax:<br>`mFile.NameRemove file, section, valuename` |
| Sections       | Returns the sections if a _PrivateProfile File_ as Dictionary with the _Section_ as the key and the _Values_ (see below) as the item |
| SectionExists  | Returns True when a given section [.....] exists in the file. |
|                | Syntax:<br>`If mFile.SectionExists(file, section) Then ...`|
| SectionsCopy   | Sub: Uses Sections Get/Let to copy named - or when omitted all - sections from a source to a target file, with the target file sections optionally replaced, by default is merged. When all sections are copied (i.e. no section names are provided), the option replace is used, and the target file is identical with the source file the sections will only be reorganized in ascending order. |
|                | Syntax:<br>`mFile.SectionsCopy source-file, target-file[, sections][, replace]`|
| SectionsRemove | Sub: Removes the named or all sections [.....] from a file |
|                | Syntax:<br>`mFile.SectionsRemove file, section`
| Value          | Get: Returns the value for a given name from a given section from a _PrivateProfile File_ <br>Syntax: `v = mFile.Value(file, section, name)`|
|                | Let: write a value with a given name under a given section in a _PrivateProfile File_ <br>Syntax: `mFile.Value(file, section, name) = value`|
| Values         | Returns the values from a given _PrivateProfile File_ under a given _Section_ as Dictionary with the value name as the key and the value as the item.<br>Syntax: `Set dct = mFile.Values(file, section)` |
   

## Installation
Download and import [mFile.bas][1] to your VB project.

## Usage
See table above

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-File-Services/master/source/mFile.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Directory-Services/master/source/mDct.bas
