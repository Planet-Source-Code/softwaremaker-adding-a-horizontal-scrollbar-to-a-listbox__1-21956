<div align="center">

## Adding a Horizontal ScrollBar to a ListBox


</div>

### Description

Just a simple code snippet that teaches users how to add that illusive Horizontal ScrollBar to a Listbox Control or any other control, for that matter.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SoftwareMaker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/softwaremaker.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/softwaremaker-adding-a-horizontal-scrollbar-to-a-listbox__1-21956/archive/master.zip)





### Source Code

```
'Declaring the SendMessage API - To send a Message to other Windows
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194
'Set the Horizontal Bar to 2 times its Width
Dim lngReturn As Long
Dim lngExtent As Long
 lngExtent = 2 * (Form1.List1.Width / Screen.TwipsPerPixelX)
 lngReturn = SendMessage(Form1.List1.hWnd, LB_SETHORIZONTALEXTENT, _
 lngExtent, 0&)
```

