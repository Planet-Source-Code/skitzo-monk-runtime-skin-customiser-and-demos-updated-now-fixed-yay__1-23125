Attribute VB_Name = "modpaint3"
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Type POINTAPI
   x As Double
   y As Double
End Type

Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Dim X1, Y1

Public Function loadfrmnoise()
frmNoise.Show
Call picblt
End Function

Public Function picblt()
BitBlt frmNoise.Picture1.hdc, 0, 0, frmNoise.Picture1.Width, frmNoise.Picture1.Height, frmpaint.picBoard.hdc, X1, Y1, SRCAND
End Function

Public Function noise()


frmNoise.Picture1.Cls
Call picblt
frmNoise.Label2.Caption = frmNoise.Slider1.Value
Randomize
For i = 1 To frmNoise.Slider1.Value
d = Int(Rnd * 2) + 1
If d = 1 Then
c = vbBlack
Else
c = vbWhite
End If
c1 = Int(Rnd * 255) + 1
c2 = Int(Rnd * 255) + 1
c3 = Int(Rnd * 255) + 1
x = Int(Rnd * frmNoise.Picture1.ScaleWidth) + 1
y = Int(Rnd * frmNoise.Picture1.ScaleHeight) + 1
If frmNoise.Option1.Value = True Then
SetPixel frmNoise.Picture1.hdc, x, y, c
Else
SetPixel frmNoise.Picture1.hdc, x, y, RGB(c1, c2, c3)
End If

Next

End Function

Public Function noiseOk()
Randomize
For i = 1 To frmNoise.Slider1.Value
d = Int(Rnd * 2) + 1
If d = 1 Then
c = vbBlack
Else
c = vbWhite
End If
c1 = Int(Rnd * 255) + 1
c2 = Int(Rnd * 255) + 1
c3 = Int(Rnd * 255) + 1
x = Int(Rnd * frmpaint.picBoard.Width) + 1
y = Int(Rnd * frmpaint.picBoard.Height) + 1
If frmNoise.Option1.Value = True Then
SetPixel frmpaint.picBoard.hdc, x, y, c
Else
SetPixel frmpaint.picBoard.hdc, x, y, RGB(c1, c2, c3)
End If

Next
End Function

Public Function mousemove(x As Single, y As Single)
X1 = x
Y1 = y
If x < 0 Then
x = 0
ElseIf y < 0 Then
y = 0
ElseIf x > frmNoise.Picture1.Width Then
x = frmNoise.Picture1.Width
ElseIf y > frmNoise.Picture1.Height Then
y = frmNoise.Picture1.Height
End If

frmNoise.Picture1.Cls
 BitBlt frmNoise.Picture1.hdc, 0, 0, frmNoise.Picture1.Width, frmNoise.Picture1.Height, frmpaint.picBoard.hdc, x, y, SRCAND

Exit Function
End Function





