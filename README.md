<div align="center">

## Passing a control array


</div>

### Description

Working with control arrays in VB3 was frustrating, but with VB4 you can pass a control array as an argument to a function. Simply specify the parameter type as Variant:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-pro.md)
**Level**          |Unknown
**User Rating**    |4.3 (171 globes from 40 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-pro-passing-a-control-array__1-103/archive/master.zip)





### Source Code

```
Private Sub Command1_Click(Index As Integer)
GetControls Command1()
End Sub
Public Sub GetControls(CArray As Variant)
Dim C As Control
For Each C In CArray
MsgBox C.Index
Next
End Sub
Also, VB4's control arrays have LBound, Ubound, and Count properties:
If Command1.Count < Command1.Ubound - _
Command1.Lbound + 1 Then _
Msgbox "Array not contiguous"
```

