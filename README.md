<div align="center">

## Select List Box or Combo Box Value by Index


</div>

### Description

This function will select the value of a List Box or Combo Box based upon the Index ID. This is helpful when you are trying to edit a record and want to select a saved value in a combo box or list box.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cierra Computers & Consulting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cierra-computers-consulting.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cierra-computers-consulting-select-list-box-or-combo-box-value-by-index__1-6054/archive/master.zip)

### API Declarations

```
Public Enum CtlType
 ListBox
 ComboBox
End Enum
```


### Source Code

```
Public Sub SelectInList(varID As Variant, ctlList As Control, Optional ctl As CtlType, _
   Optional blnRefresh As Boolean = True)
'Selects the Item in List or Combo Box that matches passed varID
Dim x
If Not IsNull(varID) Then
   varID = CLng(varID)
   If blnRefresh = True Then
     ctlList.Refresh
   End If
   For x = 0 To ctlList.ListCount - 1
     If ctlList.ItemData(x) = varID Then
        If ctl = ListBox Then
          ctlList.Selected(x) = True
        Else
          ctlList = ctlList.List(x)
        End If
        Exit Sub
     End If
   Next
Else
   'Reset the ComboBox
   ctlList.ListIndex = -1
End If
End Sub
```

