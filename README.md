<div align="center">

## Clear all Texbox in a Form


</div>

### Description

This Code clears all textboxes in a form.This code is useful if u use so much textbox contols in a form and have to clear all of them when the user selects (New Record) selection.
 
### More Info
 
no side effects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LEVENT YILDIZ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/levent-yildiz.md)
**Level**          |Advanced
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/levent-yildiz-clear-all-texbox-in-a-form__1-9028/archive/master.zip)





### Source Code

```
For i = 1 To Me.Controls.Count - 1
    If TypeOf Me.Controls(i) Is TextBox Then
      Me.Controls(i).Text = ""
    End If
  Next i
```

