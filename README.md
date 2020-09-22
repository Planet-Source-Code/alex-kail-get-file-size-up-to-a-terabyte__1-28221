<div align="center">

## Get file size up to a terabyte


</div>

### Description

Returns the file size of the file name passed in the format the user specifies
 
### More Info
 
File name, and return type (bytes, kilobytes, etc)

File size (double) in the format specified in the arguments

None to my knowledge


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex Kail](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-kail.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-kail-get-file-size-up-to-a-terabyte__1-28221/archive/master.zip)

### API Declarations

```
'Various views for file sizes
Public Enum FileSizeView
  fsBytes = -1
  fsKilobytes = 0
  fsMegabytes = 1
  fsGigabytes = 2
  fsTerabytes = 3
End Enum
```


### Source Code

```
Function FileSize(ByVal strFile As String, Optional ByVal ReturnAs As FileSizeView = fsBytes) As Double
 'Purpose: Returns the file size of the file name passed in the format the user specifies
 Dim dblLen As Double, lngIndex As Long
 'If file doesn't exist, abort
 If Dir(strFile) = Empty Then
  FileSize = 0
  Exit Function 'Abort
 End If
 'Returns the file length in bytes
 dblLen = FileLen(strFile)
 'Calculate to the file size view passed
 For lngIndex = 0 To ReturnAs
  dblLen = dblLen / 1024
 Next
 'Return the file size
 FileSize = dblLen
End Function
```

