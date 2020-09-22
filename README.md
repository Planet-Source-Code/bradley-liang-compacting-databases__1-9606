<div align="center">

## Compacting Databases


</div>

### Description

The problem with Access databases is that when you delete records, the .MDB file doesn't shrink.

It just grows and grows and grows &#8211; until someone either compacts it or you run out of disk space.

This tip will show you how to compact a JET database up to 100 times!
 
### More Info
 
Simply run CompactDatabase passing the location of your database. There's also an optional argument requiring a True or False value to backup the original database to the Temp directory before proceeding.

Note: In order for this to work, you need a reference (Project, References) to any version of the Microsoft DAO object library.

Substantially smaller Database (e.g. 25.3 mb to 4.7 mb).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bradley Liang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bradley-liang.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bradley-liang-compacting-databases__1-9606/archive/master.zip)

### API Declarations

```
Public Declare Function GetTempPath Lib "kernel32" Alias _
 "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
 As String) As Long
Public Const MAX_PATH = 260
```


### Source Code

```
Public Sub CompactDatabase(Location As String, _
 Optional BackupOriginal As Boolean = True)
On Error GoTo CompactErr
Dim strBackupFile As String
Dim strTempFile As String
'Check the database exists
If Len(Dir(Location)) Then
	' Create Backup
	If BackupOriginal = True Then
		strBackupFile = GetTemporaryPath & "backup.mdb"
		If Len(Dir(strBackupFile)) Then Kill strBackupFile
		FileCopy Location, strBackupFile
	End If
	strTempFile = GetTemporaryPath & "temp.mdb"
	If Len(Dir(strTempFile)) Then Kill strTempFile
	' Do the compacting
  'DBEngine is a reference to the Microsoft DAO Object Lib...
	DBEngine.CompactDatabase Location, strTempFile
	' Remove the uncompressed database
	Kill Location
	' Replace Uncompressed
	FileCopy strTempFile, Location
	Kill strTempFile
End If
CompactErr:
 Exit Sub
End Sub
Public Function GetTemporaryPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetTempPath(MAX_PATH, strFolder)
If lngResult <> 0 Then
 GetTemporaryPath = Left(strFolder, InStr(strFolder, _
	Chr(0)) - 1)
Else
 GetTemporaryPath = ""
End If
End Function
```

