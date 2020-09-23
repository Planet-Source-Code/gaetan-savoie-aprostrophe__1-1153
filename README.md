<div align="center">

## Aprostrophe


</div>

### Description

Have you ever try so send a string variable to MS Access that have

apostrophes using a SQL Statement? If YES you will get a run time ERROR

Here is your solution....A function that formats the

variable before sending it to the database.
 
### More Info
 
sFieldString

This code should be used in your Classes.

For example :

let say myVar=" Gaetan's"

the follwing statement will give you errors:

SSQL="INSERT INTO tablename (FirstName) VALUES (" & chr(39) & myvar & chr(39) & ")"

To fix it do the following:

myVar=apostrophe(myvar)

SSQL="INSERT INTO tablename (FirstName) VALUES (" & chr(39) & myvar & chr(39) & ")"

Aphostrophe


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gaetan Savoie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gaetan-savoie.md)
**Level**          |Unknown
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gaetan-savoie-aprostrophe__1-1153/archive/master.zip)

### API Declarations

None


### Source Code

```
'***********************************************************************
' Function: Apostrophe
' Argument: sFieldString
' Description: This subroutine will fill format the field we
' want to store in the database if there is some apostrophes
' in the field.
'***********************************************************************
Public Function Apostrophe(sFieldString As String) As String
If InStr(sFieldString, "'") Then
  Dim iLen As Integer
  Dim ii As Integer
  Dim apostr As Integer
  iLen = Len(sFieldString)
  ii = 1
  Do While ii <= iLen
   If Mid(sFieldString, ii, 1) = "'" Then
   apostr = ii
sFieldString = Left(sFieldString, apostr) & "'" & _
Right(sFieldString, iLen - apostr)
   iLen = Len(sFieldString)
   ii = ii + 1
   End If
   ii = ii + 1
  Loop
End If
Apostrophe = sFieldString
End Function
```

