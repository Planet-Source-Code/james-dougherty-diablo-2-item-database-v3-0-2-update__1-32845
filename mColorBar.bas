Attribute VB_Name = "mColorBar"
Option Explicit

Private Function GetRed(ColorVal As Long) As Integer
 GetRed = ColorVal Mod 256
End Function

Private Function GetGreen(ColorVal As Long) As Integer
 GetGreen = ((ColorVal And &HFF00FF00) / 256&)
End Function

Private Function GetBlue(ColorVal As Long) As Integer
 GetBlue = (ColorVal And &HFF0000) / 65536
End Function

Public Sub Draw_Gradient_Title_Bar_Horizontal(Picture As PictureBox, StartColor As Long, EndColor As Long)
On Local Error Resume Next
Dim NewColor As Long
Dim ipixel As Integer, PWidth As Integer
Dim RedInc As Single, GreenInc As Single, BlueInc As Single
Dim Color1 As Long: Dim Color2 As Long
Dim StartRed As Integer, StartGreen As Integer, StartBlue As Integer
Dim EndRed As Integer, EndGreen As Integer, EndBlue As Integer

Color1 = StartColor
Color2 = EndColor
    
StartRed = GetRed(Color1)
EndRed = GetRed(Color2)
StartGreen = GetGreen(Color1)
EndGreen = GetGreen(Color2)
StartBlue = GetBlue(Color1)
EndBlue = GetBlue(Color2)

PWidth = Picture.ScaleWidth
RedInc = (EndRed - StartRed) / PWidth
GreenInc = (EndGreen - StartGreen) / PWidth
BlueInc = (EndBlue - StartBlue) / PWidth
    
For ipixel = 0 To PWidth - 1
 NewColor = RGB(StartRed + RedInc * ipixel, StartGreen + GreenInc * ipixel, StartBlue + BlueInc * ipixel)
 Picture.Line (ipixel, 0)-(ipixel, Picture.Height - 1), NewColor
Next

End Sub
