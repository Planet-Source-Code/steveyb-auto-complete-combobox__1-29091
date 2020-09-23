<div align="center">

## Auto\-complete combobox


</div>

### Description

The following code adds the autocomplete functionality to the standard combo box.

The user can either type into the combobox, which automatically selects an item matching the users keypresses, or they can simply select an item in the standard fashion.
 
### More Info
 
A combo box, named Combo1, needs to be placed on your form, with the Sorted property set to TRUE.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SteveyB](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steveyb.md)
**Level**          |Intermediate
**User Rating**    |4.8 (81 globes from 17 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steveyb-auto-complete-combobox__1-29091/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Form_Load()
  With Combo1
   .AddItem "Dog"
   .AddItem "Growl"
   .AddItem "Sausage"
   .AddItem "Woof"
   .Text = ""
  End With
End Sub
Private Sub Combo1_KeyDown(keycode As Integer, Shift As Integer)
  If keycode = vbKeyDelete Then
   Combo1.Text = ""
   keycode = 0
  End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim strSearchText  As String
Dim strEnteredText As String
Dim intLength    As Integer
Dim intIndex    As Integer
Dim intCounter   As Integer
On Error GoTo ErrorHandler
  With Combo1
   If .SelStart > 0 Then
     strEnteredText = Left(.Text, .SelStart)
   End If
   Select Case KeyAscii
     Case vbKeyReturn
      If .ListIndex > -1 Then
        .SelStart = 0
        .SelLength = Len(.List(.ListIndex))
        Exit Sub
      End If
     Case vbKeyEscape, vbKeyDelete
      .Text = ""
       KeyAscii = 0
       Exit Sub
     Case vbKeyBack
       If Len(strEnteredText) > 1 Then
        strSearchText = LCase(Left(strEnteredText, Len(strEnteredText) - 1))
       Else
        strEnteredText = ""
        KeyAscii = 0
        .Text = ""
        Exit Sub
       End If
     Case Else
      strSearchText = LCase(strEnteredText & Chr(KeyAscii))
   End Select
   intIndex = -1
   intLength = Len(strSearchText)
   For intCounter = 0 To .ListCount - 1
     If LCase(Left(.List(intCounter), intLength)) = strSearchText Then
      intIndex = intCounter
      Exit For
     End If
   Next intCounter
   If intIndex > -1 Then
     .ListIndex = intIndex
     .SelStart = Len(strSearchText)
     .SelLength = Len(.List(intIndex)) - Len(strSearchText)
   Else
     Beep
   End If
  End With
  KeyAscii = 0
  Exit Sub
ErrorHandler:
  KeyAscii = 0
  Beep
End Sub
Private Sub Combo1_LostFocus()
  Combo1.SelLength = 0
End Sub
```

