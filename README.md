<div align="center">

## Auto Save/Load/Add/Remove \- Lists


</div>

### Description

Auto Save/Load/Add/Remove from/to a List
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Gates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-gates.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-gates-auto-save-load-add-remove-lists__1-9010/archive/master.zip)





### Source Code

```
Public Sub List_Add(List As ListBox, txt As String)
List.AddItem txt
End Sub
Public Sub List_Load(TheList As ListBox, FileName As String)
'Loads a file to a list box
On Error Resume Next
Dim TheContents As String
Dim fFile As Integer
fFile = FreeFile
 Open FileName For Input As fFile
  Do
   Line Input #fFile, TheContents$
    Call List_Add(TheList, TheContents$)
  Loop Until EOF(fFile)
 Close fFile
End Sub
Public Sub List_Save(TheList As ListBox, FileName As String)
'Save a listbox as FileName
On Error Resume Next
Dim Save As Long
Dim fFile As Integer
fFile = FreeFile
Open FileName For Output As fFile
  For Save = 0 To TheList.ListCount - 1
   Print #fFile, TheList.List(Save)
  Next Save
Close fFile
End Sub
Public Sub List_Remove(List As ListBox)
On Error Resume Next
If List.ListCount < 0 Then Exit Sub
 List.RemoveItem List.ListIndex
End Sub
```

