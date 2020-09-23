<div align="center">

## Move Form without a Title Bar\!


</div>

### Description

This code will allow you to move your Forms without even having to have a Title Bar! So this means if you choose to make your Form's BroderStyle 0-None, which means no TitleBar, you will still be able to move the form with this code!

You can do multiple things with this code also like: Clicking on the form and dragging to move the form, clicking on a Label and dragging it to move the form, clicking on a CommandButton and dragging it to move the form and so on if you get the picture :-) This code is very useful and cool if your sick of that dumb old BlueBar on the top of your form and want to make your own cool TitleBars and Borders and anything else that you put your mind to!
 
### More Info
 
'Follow these steps and don't skip anything!

'1.)Start a New Project in your 32bit Visual Basics

'2.)Add a New Module/Bas to your New Project

'3.)Add a Label to the form and name it: "Label1"

'4.)Add a CommandButton to the form and name it: "Command1"

'5.)Make Form1's BorderStyle 0-None


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael L\. Canejo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-l-canejo.md)
**Level**          |Beginner
**User Rating**    |4.6 (65 globes from 14 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-l-canejo-move-form-without-a-title-bar__1-1486/archive/master.zip)

### API Declarations

```
'Type the following in the Module/Bas!! NOT IN THE FORM!! (it wont work!)
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Sub FormDrag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
```


### Source Code

```
'Copy and Paste the following below this in the Form. NOT THE MODULE/BAS!!!!
'Ok, here it is, start Copying:
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
```

