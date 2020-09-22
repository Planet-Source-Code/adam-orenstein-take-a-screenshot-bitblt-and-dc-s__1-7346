<div align="center">

## Take a screenshot \(BitBlt and DC's\)


</div>

### Description

My code here will allow you take a screen shot of the entire screen.
 
### More Info
 
Remeber when taking a screen shot the picture gets added to the destinations Image property not its image property. also don't forget to set the Autoredraw property of the destination of the image to True. Thanks!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Orenstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-orenstein.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-orenstein-take-a-screenshot-bitblt-and-dc-s__1-7346/archive/master.zip)

### API Declarations

```
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
```


### Source Code

```
Public Function CaptureScreen(PicDest As Object)
 DeskWnd& = GetDesktopWindow
 deskdc& = GetDC(DeskWnd&)
 Call BitBlt(PicDest.hDC, 0&, 0&, Screen.Width, Screen.Height, deskdc&, _
 0&, 0&, SRCCOPY)
 Call ReleaseDC(deskdc&, 0&)
 PicDest.Refresh
End Function
```

