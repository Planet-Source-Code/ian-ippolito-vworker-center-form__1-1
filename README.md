<div align="center">

## Center\_Form


</div>

### Description

To center all of your forms nicely on the screen, use this as the first line in the Form_Load event--resolution independent.

'note:call this function like this:

Center_Form Me
 
### More Info
 
frmForm--form to center. Call this function like this:Center_Form Me


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-center-form__1-1/archive/master.zip)





### Source Code

```
Sub Center_Form (frmForm As Form)
 frmForm.Left = (Screen.Width - frmForm.Width) / 2
 frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub
```

