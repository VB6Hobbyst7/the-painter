Attribute VB_Name = "PHMod2"
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Public Color&, TempCol&, MidX%, MidY%, Pct%, kk!
Public EchoX%, EchoY%, EchoNr%, EchoRed%, LineColor&()

Public Sub SelectAll()
FMain.Shape1.Visible = False
Xcor0 = 0
Xcor1 = FMain.Pic1.Width - 1
Ycor0 = 0
Ycor1 = FMain.Pic1.Height - 1
SetCoordinates
FMain.Toolbar1.Buttons(3).Enabled = False
FMain.mnuSel(1).Enabled = False
End Sub

Public Sub ReadColor(Rx1%, Ry1%, Rx2%, Ry2%)
On Error Resume Next
Screen.MousePointer = 11
FMain.Label2.Caption = lgT(310)
DoEvents
FMain.PB1.Value = 0
FMain.PB1.min = 0
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 '- 1
For Yy = Ry1 To Ry2 '- 1
Color = GetPixel(FMain.Pic1.hdc, Xx, Yy)
r(Xx, Yy) = Color Mod 256&
G(Xx, Yy) = ((Color And &HFF00) / 256&) Mod 256&
B(Xx, Yy) = (Color And &HFF0000) / 65536
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
FMain.PB1.Value = 0
End Sub

Public Sub KillComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
Dim Mask&
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
If Comp = 0 Then Mask = &HFFFF00
If Comp = 1 Then Mask = &HFF00FF
If Comp = 2 Then Mask = &HFFFF&
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, Xx, Yy, (GetPixel(FMain.Pic1.hdc, Xx, Yy) And Mask)
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub ColorComp(Rx1, Ry1, Rx2, Ry2, Rpct!, Gpct!, Bpct!)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
Rpct = 1 + (Rpct / 100)
Gpct = 1 + (Gpct / 100)
Bpct = 1 + (Bpct / 100)
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) * Rpct, G(Xx, Yy) * Gpct, B(Xx, Yy) * Bpct)
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub SwapComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
Rpct = 1 + (Rpct / 100)
Gpct = 1 + (Gpct / 100)
Bpct = 1 + (Bpct / 100)
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
If Comp = 0 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(B(Xx, Yy), G(Xx, Yy), r(Xx, Yy))
If Comp = 1 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(B(Xx, Yy), r(Xx, Yy), G(Xx, Yy))
If Comp = 2 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(G(Xx, Yy), B(Xx, Yy), r(Xx, Yy))
If Comp = 3 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(G(Xx, Yy), r(Xx, Yy), B(Xx, Yy))
If Comp = 4 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), B(Xx, Yy), G(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub PhotoNeg(Rx1%, Ry1%, Rx2%, Ry2%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) Xor 255, B(Xx, Yy) Xor 255, G(Xx, Yy) Xor 255)
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub PhotoNegComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
If Comp = 0 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) Xor 255, G(Xx, Yy), B(Xx, Yy))
If Comp = 1 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy) Xor 255, B(Xx, Yy))
If Comp = 2 Then SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy) Xor 255)
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub GreyColor(Rx1%, Ry1%, Rx2%, Ry2%)  'grey
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
    r(Xx, Yy) = r(Xx, Yy) * 0.3 + G(Xx, Yy) * 0.59 + B(Xx, Yy) * 0.11
    If r(Xx, Yy) > 255 Then r(Xx, Yy) = 255
    If r(Xx, Yy) < 0 Then r(Xx, Yy) = 0
    G(Xx, Yy) = r(Xx, Yy)
    B(Xx, Yy) = r(Xx, Yy)
    SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub ContrastPic(Rx1, Ry1, Rx2, Ry2, Rpct!)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
    If r(Xx, Yy) > 127 Then
    r(Xx, Yy) = r(Xx, Yy) + Rpct
    Else
    r(Xx, Yy) = r(Xx, Yy) - Rpct
    End If
    If G(Xx, Yy) > 127 Then
    G(Xx, Yy) = G(Xx, Yy) + Rpct
    Else
    G(Xx, Yy) = G(Xx, Yy) - Rpct
    End If
    If B(Xx, Yy) > 127 Then
    B(Xx, Yy) = B(Xx, Yy) + Rpct
    Else
    B(Xx, Yy) = B(Xx, Yy) - Rpct
    End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub FlipX() 'flip horizontal
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width, FMain.Pic1.Height
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub FlipY() 'flip vertical
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub MirrorX() 'mirror x
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width / 2, FMain.Pic1.Height, 0, 0, FMain.Pic1.Width / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub MirrorXRev() 'mirror x
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width, FMain.Pic1.Height
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width / 2, FMain.Pic1.Height, 0, 0, FMain.Pic1.Width / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub MirrorY() 'mirror y
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height / 2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub MirrorYRev() 'mirror y
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height / 2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = lgT(311)
End Sub

Public Sub EmbossPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'emboss
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = (Abs(r(Xx, Yy) - r(Xx + 1, Yy + 1) + 128))
G(Xx, Yy) = (Abs(G(Xx, Yy) - G(Xx + 1, Yy + 1) + 128))
B(Xx, Yy) = (Abs(B(Xx, Yy) - B(Xx + 1, Yy + 1) + 128))
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub NeonPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'emboss
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = (Abs(r(Xx - 1, Yy) + r(Xx, Yy) - r(Xx + 1, Yy) - r(Xx + 2, Yy) + 32))
G(Xx, Yy) = (Abs(G(Xx - 1, Yy) + G(Xx, Yy) - G(Xx + 1, Yy) - G(Xx + 2, Yy) + 32))
B(Xx, Yy) = (Abs(B(Xx - 1, Yy) + B(Xx, Yy) - B(Xx + 1, Yy) - B(Xx + 2, Yy) + 32))
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub EngravePicture(Rx1%, Ry1%, Rx2%, Ry2%) 'engrave
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = (Abs(r(Xx + 1, Yy + 1) - r(Xx, Yy) + 128))
G(Xx, Yy) = (Abs(G(Xx + 1, Yy + 1) - G(Xx, Yy) + 148))
B(Xx, Yy) = (Abs(B(Xx + 1, Yy + 1) - B(Xx, Yy) + 128))
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub BeginProcess()
ReadColor 0, 0, FMain.Pic1.Width - 1, FMain.Pic1.Height - 1
FMain.Label2.Caption = lgT(312)
DoEvents
FMain.PB1.Value = 0
FMain.PB1.min = 0
End Sub

Public Sub EndProcess()
FMain.PB1.Value = 0
FMain.Pic1.Refresh
FMain.Label2.Caption = lgT(311)
Screen.MousePointer = 1
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub HoldRed(Rx1, Ry1, Rx2, Ry2) 'Hold red
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
If r(Xx, Yy) < 128 Then
r(Xx, Yy) = (Abs(r(Xx, Yy) - r(Xx + 1, Yy + 1) + 128))
G(Xx, Yy) = (Abs(G(Xx, Yy) - G(Xx + 1, Yy + 1) + 128))
B(Xx, Yy) = (Abs(B(Xx, Yy) - B(Xx + 1, Yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub HoldGreen(Rx1, Ry1, Rx2, Ry2) 'Hold green
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
If G(Xx, Yy) < 128 Then
r(Xx, Yy) = (Abs(r(Xx, Yy) - r(Xx + 1, Yy + 1) + 128))
G(Xx, Yy) = (Abs(G(Xx, Yy) - G(Xx + 1, Yy + 1) + 128))
B(Xx, Yy) = (Abs(B(Xx, Yy) - B(Xx + 1, Yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub HoldBlue(Rx1, Ry1, Rx2, Ry2) 'Hold blue
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
If B(Xx, Yy) < 128 Then
r(Xx, Yy) = (Abs(r(Xx, Yy) - r(Xx + 1, Yy + 1) + 128))
G(Xx, Yy) = (Abs(G(Xx, Yy) - G(Xx + 1, Yy + 1) + 128))
B(Xx, Yy) = (Abs(B(Xx, Yy) - B(Xx + 1, Yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub BlurPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'blur
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 1 To Rx2 - 1
For Yy = Ry1 + 1 To Ry2 - 1
r(Xx, Yy) = (Abs(r(Xx - 1, Yy - 1) + r(Xx - 1, Yy) + r(Xx - 1, Yy + 1) + r(Xx, Yy - 1) + r(Xx, Yy) + r(Xx, Yy + 1) + r(Xx + 1, Yy - 1) + r(Xx + 1, Yy) + r(Xx + 1, Yy + 1))) / 9
G(Xx, Yy) = (Abs(G(Xx - 1, Yy - 1) + G(Xx - 1, Yy) + G(Xx - 1, Yy + 1) + G(Xx, Yy - 1) + G(Xx, Yy) + G(Xx, Yy + 1) + G(Xx + 1, Yy - 1) + G(Xx + 1, Yy) + G(Xx + 1, Yy + 1))) / 9
B(Xx, Yy) = (Abs(B(Xx - 1, Yy - 1) + B(Xx - 1, Yy) + B(Xx - 1, Yy + 1) + B(Xx, Yy - 1) + B(Xx, Yy) + B(Xx, Yy + 1) + B(Xx + 1, Yy - 1) + B(Xx + 1, Yy) + B(Xx + 1, Yy + 1))) / 9
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub BlurPictureMore(Rx1%, Ry1%, Rx2%, Ry2%)  'blur more
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 2 To Rx2 - 1
For Yy = Ry1 + 2 To Ry2 - 1
r(Xx, Yy) = (Abs(r(Xx - 2, Yy - 2) + r(Xx - 2, Yy - 1) + r(Xx - 2, Yy) + r(Xx - 2, Yy + 1) + r(Xx - 2, Yy + 2) + r(Xx - 1, Yy - 2) + r(Xx - 1, Yy - 1) + r(Xx - 1, Yy) + r(Xx - 1, Yy + 1) + r(Xx - 1, Yy + 2) + r(Xx, Yy - 2) + r(Xx, Yy - 1) + r(Xx, Yy) + r(Xx, Yy + 1) + r(Xx, Yy + 2) + r(Xx + 1, Yy - 2) + r(Xx + 1, Yy - 1) + r(Xx + 1, Yy) + r(Xx + 1, Yy + 1) + r(Xx + 1, Yy + 2) + r(Xx + 2, Yy - 2) + r(Xx + 2, Yy - 1) + r(Xx + 2, Yy) + r(Xx + 2, Yy + 1) + r(Xx + 2, Yy + 2))) / 25
G(Xx, Yy) = (Abs(G(Xx - 2, Yy - 2) + G(Xx - 2, Yy - 1) + G(Xx - 2, Yy) + G(Xx - 2, Yy + 1) + G(Xx - 2, Yy + 2) + G(Xx - 1, Yy - 2) + G(Xx - 1, Yy - 1) + G(Xx - 1, Yy) + G(Xx - 1, Yy + 1) + G(Xx - 1, Yy + 2) + G(Xx, Yy - 2) + G(Xx, Yy - 1) + G(Xx, Yy) + G(Xx, Yy + 1) + G(Xx, Yy + 2) + G(Xx + 1, Yy - 2) + G(Xx + 1, Yy - 1) + G(Xx + 1, Yy) + G(Xx + 1, Yy + 1) + G(Xx + 1, Yy + 2) + G(Xx + 2, Yy - 2) + G(Xx + 2, Yy - 1) + G(Xx + 2, Yy) + G(Xx + 2, Yy + 1) + G(Xx + 2, Yy + 2))) / 25
B(Xx, Yy) = (Abs(B(Xx - 2, Yy - 2) + B(Xx - 2, Yy - 1) + B(Xx - 2, Yy) + B(Xx - 2, Yy + 1) + B(Xx - 2, Yy + 2) + B(Xx - 1, Yy - 2) + B(Xx - 1, Yy - 1) + B(Xx - 1, Yy) + B(Xx - 1, Yy + 1) + B(Xx - 1, Yy + 2) + B(Xx, Yy - 2) + B(Xx, Yy - 1) + B(Xx, Yy) + B(Xx, Yy + 1) + B(Xx, Yy + 2) + B(Xx + 1, Yy - 2) + B(Xx + 1, Yy - 1) + B(Xx + 1, Yy) + B(Xx + 1, Yy + 1) + B(Xx + 1, Yy + 2) + B(Xx + 2, Yy - 2) + B(Xx + 2, Yy - 1) + B(Xx + 2, Yy) + B(Xx + 2, Yy + 1) + B(Xx + 2, Yy + 2))) / 25
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub SharpenPicture(Rx1%, Ry1%, Rx2%, Ry2%)   'sharpen
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 1 To Rx2 - 1
For Yy = Ry1 + 1 To Ry2 - 1
r(Xx, Yy) = r(Xx, Yy) + 0.5 * (r(Xx, Yy) - r(Xx - 1, Yy - 1))
G(Xx, Yy) = G(Xx, Yy) + 0.5 * (G(Xx, Yy) - G(Xx - 1, Yy - 1))
B(Xx, Yy) = B(Xx, Yy) + 0.5 * (B(Xx, Yy) - B(Xx - 1, Yy - 1))
            If r(Xx, Yy) > 255 Then r(Xx, Yy) = 255
            If r(Xx, Yy) < 0 Then r(Xx, Yy) = 0
            If G(Xx, Yy) > 255 Then G(Xx, Yy) = 255
            If G(Xx, Yy) < 0 Then G(Xx, Yy) = 0
            If B(Xx, Yy) > 255 Then B(Xx, Yy) = 255
            If B(Xx, Yy) < 0 Then B(Xx, Yy) = 0
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub DiffusePic(Rx1%, Ry1%, Rx2%, Ry2%, Diffuse%) 'diffuse
Dim tt%, tt1%
On Error Resume Next
tt = Diffuse * 10
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
tt1 = (Rnd * tt) - 2
r(Xx, Yy) = Abs(r(Xx, Yy) + tt1)
G(Xx, Yy) = Abs(G(Xx, Yy) + tt1)
B(Xx, Yy) = Abs(B(Xx, Yy) + tt1)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub ErodePic(Rx1%, Ry1%, Rx2%, Ry2%, Erode%) 'erode
On Error Resume Next
Pct = Erode * 8
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(r(Xx, Yy) Xor Pct)
G(Xx, Yy) = Abs(G(Xx, Yy) Xor Pct)
B(Xx, Yy) = Abs(B(Xx, Yy) Xor Pct)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub BlowPic(Rx1%, Ry1%, Rx2%, Ry2%, Blow%) 'blow
On Error Resume Next
Pct = Blow
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(r(Xx, Yy) Xor (r(Xx, Yy) / Pct))
G(Xx, Yy) = Abs(G(Xx, Yy) Xor (G(Xx, Yy) / Pct))
B(Xx, Yy) = Abs(B(Xx, Yy) Xor (B(Xx, Yy) / Pct))
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub AddNoise(Rx1%, Ry1%, Rx2%, Ry2%) 'addnoise
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 1 To Rx2 - 1
For Yy = Ry1 + 1 To Ry2 - 1
r(Xx, Yy) = ((Rnd * r(Xx, Yy)) + r(Xx, Yy)) / 2
G(Xx, Yy) = ((Rnd * G(Xx, Yy)) + G(Xx, Yy)) / 2
B(Xx, Yy) = ((Rnd * B(Xx, Yy)) + B(Xx, Yy)) / 2
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub FogPic(Rx1%, Ry1%, Rx2%, Ry2%, Fog%) 'fog
Dim tt1%
On Error Resume Next
Pct = Fog
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Yy = Ry1 To Ry2 - 1
tt1 = (Rnd * Pct) - 2
For Xx = Rx1 To Rx2 - 1
r(Xx, Yy) = Abs(r(Xx, Yy) + tt1)
G(Xx, Yy) = Abs(G(Xx, Yy) + tt1)
B(Xx, Yy) = Abs(B(Xx, Yy) + tt1)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Xx
FMain.PB1.Value = Yy - Ry1
Next Yy
EndProcess
End Sub

Public Sub FreezePic(Rx1%, Ry1%, Rx2%, Ry2%, Freeze!) 'freeze
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 1 To Rx2 - 1
For Yy = Ry1 + 1 To Ry2 - 1
r(Xx, Yy) = Abs((r(Xx, Yy) - G(Xx, Yy) - B(Xx, Yy)) * Freeze)
G(Xx, Yy) = Abs((G(Xx, Yy) - B(Xx, Yy) - r(Xx, Yy)) * Freeze)
B(Xx, Yy) = Abs((B(Xx, Yy) - r(Xx, Yy) - G(Xx, Yy)) * Freeze)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub BnW(Rx1%, Ry1%, Rx2%, Ry2%, BW%) 'B & W
Dim BWColor&
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 + 1 To Rx2 - 1
For Yy = Ry1 + 1 To Ry2 - 1
    If r(Xx, Yy) < BW And G(Xx, Yy) < BW And B(Xx, Yy) < BW Then
    BWColor = 0
    Else
    BWColor = &HFFFFFF
    End If
SetPixel FMain.Pic1.hdc, Xx, Yy, BWColor
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Effect0(Rx1%, Ry1%, Rx2%, Ry2%, Eff%)
On Error Resume Next
Dim C&
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Yy = Ry1 To Ry2 - 1
For Xx = Rx1 To Rx2 - 1
Select Case Eff
Case 0
    G(Xx, Yy) = (r(Xx, Yy) + G(Xx, Yy)) / 2
    r(Xx, Yy) = G(Xx, Yy)
    B(Xx, Yy) = B(Xx, Yy) * (Atn(G(Xx, Yy)) * 2)
Case 1
    G(Xx, Yy) = (r(Xx, Yy) + G(Xx, Yy)) / 2
    r(Xx, Yy) = G(Xx, Yy)
Case 2
    r(Xx, Yy) = (r(Xx, Yy) + B(Xx, Yy)) / 2
    B(Xx, Yy) = r(Xx, Yy)
Case 3
    G(Xx, Yy) = (G(Xx, Yy) + B(Xx, Yy)) / 2
    B(Xx, Yy) = G(Xx, Yy)
Case 4
    G(Xx, Yy) = (B(Xx, Yy) + G(Xx, Yy)) / 2
    B(Xx, Yy) = G(Xx, Yy)
    r(Xx, Yy) = r(Xx, Yy) * (Atn(G(Xx, Yy)) * 2)
Case 5
    B(Xx, Yy) = (B(Xx, Yy) + r(Xx, Yy)) / 2
    r(Xx, Yy) = B(Xx, Yy)
    G(Xx, Yy) = G(Xx, Yy) * (Atn(r(Xx, Yy)) * 2)
Case 6
    B(Xx, Yy) = Sin(B(Xx, Yy)) * B(Xx, Yy)
    r(Xx, Yy) = Sin(r(Xx, Yy)) * r(Xx, Yy)
    G(Xx, Yy) = Sin(G(Xx, Yy)) * G(Xx, Yy)
Case 7
    C = (r(Xx, Yy) + G(Xx, Yy) + B(Xx, Yy)) / 12
    B(Xx, Yy) = Abs(Not (G(Xx, Yy) + C))
    r(Xx, Yy) = Abs(Not (B(Xx, Yy) + C))
    G(Xx, Yy) = Abs(Not (r(Xx, Yy) + C))
Case 8
    B(Xx, Yy) = G(Xx, Yy)
    G(Xx, Yy) = r(Xx, Yy)
Case 9
    r(Xx, Yy) = r(Xx, Yy) / 2
    B(Xx, Yy) = G(Xx, Yy) / 2
    G(Xx, Yy) = r(Xx, Yy)
Case 10
    r(Xx, Yy) = r(Xx, Yy)
    B(Xx, Yy) = G(Xx, Yy) / 2
    G(Xx, Yy) = r(Xx, Yy) / 2
Case 11
    r(Xx, Yy) = r(Xx, Yy) + Abs(Sin(r(Xx, Yy)) * 64)
    G(Xx, Yy) = G(Xx, Yy) + Abs(Sin(G(Xx, Yy)) * 64)
    B(Xx, Yy) = B(Xx, Yy) + Abs(Sin(B(Xx, Yy)) * 64)
End Select
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Xx
FMain.PB1.Value = Yy - Ry1
Next Yy
EndProcess
End Sub

Public Sub Brown(Rx1%, Ry1%, Rx2%, Ry2%, Brown%) 'brown
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(G(Xx, Yy) * B(Xx, Yy)) / Brown
G(Xx, Yy) = Abs(B(Xx, Yy) * r(Xx, Yy)) / 256
B(Xx, Yy) = Abs(r(Xx, Yy) * G(Xx, Yy)) / 256
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Liquid(Rx1%, Ry1%, Rx2%, Ry2%) 'liquid
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = ((G(Xx, Yy) - B(Xx, Yy)) ^ 2) / 125
G(Xx, Yy) = ((r(Xx, Yy) - B(Xx, Yy)) ^ 2) / 125
B(Xx, Yy) = ((r(Xx, Yy) - G(Xx, Yy)) ^ 2) / 125
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Yellow(Rx1%, Ry1%, Rx2%, Ry2%) 'yellow
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
B(Xx, Yy) = ((G(Xx, Yy) - r(Xx, Yy)) ^ 2) / 125
r(Xx, Yy) = ((G(Xx, Yy) - B(Xx, Yy)) ^ 2) / 125
G(Xx, Yy) = ((B(Xx, Yy) + r(Xx, Yy)) ^ 2) / 125
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Charcoal(Rx1%, Ry1%, Rx2%, Ry2%) 'charcoal
Dim tCol&
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
            r(Xx, Yy) = Abs(r(Xx, Yy) * (G(Xx, Yy) - B(Xx, Yy) + G(Xx, Yy) + r(Xx, Yy))) / 256
            G(Xx, Yy) = Abs(r(Xx, Yy) * (B(Xx, Yy) - G(Xx, Yy) + B(Xx, Yy) + r(Xx, Yy))) / 256
            B(Xx, Yy) = Abs(G(Xx, Yy) * (B(Xx, Yy) - G(Xx, Yy) + B(Xx, Yy) + r(Xx, Yy))) / 256
            tCol = RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
            r(Xx, Yy) = Abs(tCol Mod 256)
            G(Xx, Yy) = Abs((tCol \ 256) Mod 256)
            B(Xx, Yy) = Abs(tCol \ 256 \ 256)
            r(Xx, Yy) = (r(Xx, Yy) + G(Xx, Yy) + B(Xx, Yy)) / 3
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), r(Xx, Yy), r(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub DarkMoon(Rx1%, Ry1%, Rx2%, Ry2%) 'dark moon
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(r(Xx, Yy) - 64)
G(Xx, Yy) = Abs(r(Xx, Yy) - 64)
B(Xx, Yy) = Abs(r(Xx, Yy) - 64)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub TotalEclipse(Rx1%, Ry1%, Rx2%, Ry2%) 'eclipse
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(G(Xx, Yy) - 64)
G(Xx, Yy) = Abs(G(Xx, Yy) - 64)
B(Xx, Yy) = Abs(G(Xx, Yy) - 64)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub PurpleRain(Rx1%, Ry1%, Rx2%, Ry2%) 'purple
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = Abs(G(Xx, Yy) + r(Xx, Yy) / 2)
G(Xx, Yy) = Abs(B(Xx, Yy) + G(Xx, Yy) / 2)
B(Xx, Yy) = Abs(r(Xx, Yy) + B(Xx, Yy) / 2)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Spooky(Rx1%, Ry1%, Rx2%, Ry2%) 'Spooky
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
G(Xx, Yy) = Abs(r(Xx, Yy) + G(Xx, Yy) / 2)
B(Xx, Yy) = Abs(r(Xx, Yy) + B(Xx, Yy) / 2)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub UnReal(Rx1%, Ry1%, Rx2%, Ry2%) 'unreal
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
If (G(Xx, Yy) = 0) Or (B(Xx, Yy) = 0) Then
    G(Xx, Yy) = 1
    B(Xx, Yy) = 1
End If
        r(Xx, Yy) = Abs(Sin(Atn(G(Xx, Yy) / B(Xx, Yy))) * 125 + 20)
        G(Xx, Yy) = Abs(Sin(Atn(r(Xx, Yy) / B(Xx, Yy))) * 125 + 20)
        B(Xx, Yy) = Abs(Sin(Atn(r(Xx, Yy) / G(Xx, Yy))) * 125 + 20)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Flame(Rx1%, Ry1%, Rx2%, Ry2%) 'flame
Dim C As Long
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
    C = (r(Xx, Yy) + G(Xx, Yy) + B(Xx, Yy)) / 3
        If r(Xx, Yy) > B(Xx, Yy) Then
            r(Xx, Yy) = Abs(r(Xx, Yy) + C)
            B(Xx, Yy) = Abs(B(Xx, Yy) - C)
        End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Aquarel(Rx1%, Ry1%, Rx2%, Ry2%) 'aquarel
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
If r(Xx, Yy) < 128 And G(Xx, Yy) < 128 And B(Xx, Yy) < 128 Then
r(Xx, Yy) = 2 * r(Xx, Yy): G(Xx, Yy) = 2 * G(Xx, Yy): B(Xx, Yy) = 2 * B(Xx, Yy)
Else
r(Xx, Yy) = r(Xx, Yy) / 2: G(Xx, Yy) = G(Xx, Yy) / 2: B(Xx, Yy) = B(Xx, Yy) / 2
End If
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Blinds(Rx1, Ry1, Rx2, Ry2, Blinds%, Reverse As Boolean) 'hor blinds
Dim rt%
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Ry2 - Ry1
If Reverse = False Then
rt = 0
Else
rt = Pct
End If
For Yy = Ry1 To Ry2 - 1
For Xx = Rx1 To Rx2 - 1
r(Xx, Yy) = r(Xx, Yy) - (rt * r(Xx, Yy) / Pct)
G(Xx, Yy) = G(Xx, Yy) - (rt * G(Xx, Yy) / Pct)
B(Xx, Yy) = B(Xx, Yy) - (rt * B(Xx, Yy) / Pct)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Xx
If Reverse = False Then
    rt = rt + 1
    If rt = Pct Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Pct
End If
FMain.PB1.Value = Yy - Ry1
Next Yy
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Blinds2(Rx1, Ry1, Rx2, Ry2, Blinds%, Reverse As Boolean) 'vert blinds
Dim rt%
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Rx2 - Rx1
If Reverse = False Then
rt = 0
Else
rt = Pct
End If
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = r(Xx, Yy) - (rt * r(Xx, Yy) / Pct)
G(Xx, Yy) = G(Xx, Yy) - (rt * G(Xx, Yy) / Pct)
B(Xx, Yy) = B(Xx, Yy) - (rt * B(Xx, Yy) / Pct)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
If Reverse = False Then
    rt = rt + 1
    If rt = Pct Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Pct
End If
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Blinds3(Rx1, Ry1, Rx2, Ry2, Blinds%) 'hor bump blinds
Dim rt%, Rtt As Boolean
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Ry2 - Ry1
rt = 0
Rtt = False
For Yy = Ry1 To Ry2 - 1
For Xx = Rx1 To Rx2 - 1
r(Xx, Yy) = r(Xx, Yy) - (rt * r(Xx, Yy) / Pct)
G(Xx, Yy) = G(Xx, Yy) - (rt * G(Xx, Yy) / Pct)
B(Xx, Yy) = B(Xx, Yy) - (rt * B(Xx, Yy) / Pct)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Xx
    If Rtt = False Then
    rt = rt + 2
    Else
    rt = rt - 2
    End If
        If rt >= Pct Then Rtt = True
        If rt <= 0 Then Rtt = False
FMain.PB1.Value = Yy - Ry1
Next Yy
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Blinds4(Rx1, Ry1, Rx2, Ry2, Blinds%) 'bump vert blinds
Dim rt%, Rtt As Boolean
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Rx2 - Rx1
rt = 0
Rtt = False
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = r(Xx, Yy) - (rt * r(Xx, Yy) / Pct)
G(Xx, Yy) = G(Xx, Yy) - (rt * G(Xx, Yy) / Pct)
B(Xx, Yy) = B(Xx, Yy) - (rt * B(Xx, Yy) / Pct)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
    If Rtt = False Then
    rt = rt + 2
    Else
    rt = rt - 2
    End If
        If rt >= Pct Then Rtt = True
        If rt <= 0 Then Rtt = False
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub HLines(Rx1, Ry1, Rx2, Ry2, Dist%, AB!, LCol&)
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
BeginProcess
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1 Step Dist
r(Xx, Yy) = (r(Xx, Yy) * (1 - AB)) + (Lr * AB)
G(Xx, Yy) = (G(Xx, Yy) * (1 - AB)) + (Lg * AB)
B(Xx, Yy) = (B(Xx, Yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub VLines(Rx1, Ry1, Rx2, Ry2, Dist%, AB!, LCol&)
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
BeginProcess
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1 Step Dist
For Yy = Ry1 To Ry2 - 1
r(Xx, Yy) = (r(Xx, Yy) * (1 - AB)) + (Lr * AB)
G(Xx, Yy) = (G(Xx, Yy) * (1 - AB)) + (Lg * AB)
B(Xx, Yy) = (B(Xx, Yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub Squares(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
If Lr = 0 And Lg = 0 And Lb = 0 Then Lr = 1 'not black!
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Yy = Ry1 To Ry2 - 1 Step Dist
FMain.TempMem.Line (Rx1, Yy)-(Rx2, Yy), LCol
Next Yy
For Xx = Rx1 To Rx2 - 1 Step Dist
FMain.TempMem.Line (Xx, Ry1)-(Xx, Ry2), LCol
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddBoxes(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&) 'add boxes
On Error Resume Next
Dim ttt%
ttt = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = 0 To (FMain.Pic1.Width / (Dist * 2))
FMain.TempMem.Line (ttt, ttt)-(FMain.Pic1.Width - ttt, FMain.Pic1.Height - ttt), LCol, B
If FMain.Pic1.Width - ttt - ttt < Dist Then Exit For
ttt = ttt + Dist
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddCircles(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&) 'add circles
On Error Resume Next
Dim ttt%
ttt = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = 0 To (FMain.Pic1.Width * 2) / Dist
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), ttt, LCol
If ttt > Int(Sqr(2 * ((FMain.Pic1.Width / 2) ^ 2))) Then Exit For
ttt = ttt + Dist
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddDiaRLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, ttt)-(ttt, 0), LCol
ttt = ttt + Dist
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddDiaLLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, FMain.Pic1.Height - ttt)-(FMain.Pic1.Width, (2 * FMain.Pic1.Height) - ttt), LCol
ttt = ttt + Dist
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddCrossLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, ttt)-(ttt, 0), LCol
ttt = ttt + Dist
Next Xx
ttt = 0
For Xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, FMain.Pic1.Height - ttt)-(FMain.Pic1.Width, (2 * FMain.Pic1.Height) - ttt), LCol
ttt = ttt + Dist
Next Xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub MixObject(Rx1%, Ry1%, Rx2%, Ry2%, AB!, LCol&) 'mix with object
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1
    Color = GetPixel(FMain.TempMem.hdc, Xx, Yy)
    If Color <> 0 Then
    Lr = Color Mod 256&
    Lg = ((Color And &HFF00) / 256&) Mod 256&
    Lb = (Color And &HFF0000) / 65536
    r(Xx, Yy) = (r(Xx, Yy) * (1 - AB)) + (Lr * AB)
    G(Xx, Yy) = (G(Xx, Yy) * (1 - AB)) + (Lg * AB)
    B(Xx, Yy) = (B(Xx, Yy) * (1 - AB)) + (Lb * AB)
    SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
    End If
    Next Yy
    FMain.PB1.Value = Xx - Rx1
    Next Xx
End Sub

Public Sub SinusLineX(Rx1%, Ry1%, Rx2%, Ry2%, AB!, Wave%, Ampl%, LCol&, Dist%, Eff%)
On Error Resume Next
Dim Degree As Single, k!
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
For Xx = Rx1 To Rx2 - 1
For Yy = Ry1 To Ry2 - 1 Step Dist
Degree = Xx * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
Color = GetPixel(FMain.Pic1.hdc, Xx, k + Yy)
r(Xx, Yy) = Color Mod 256&
G(Xx, Yy) = ((Color And &HFF00) / 256&) Mod 256&
B(Xx, Yy) = (Color And &HFF0000) / 65536
r(Xx, Yy) = (r(Xx, Yy) * (1 - AB)) + (Lr * AB)
G(Xx, Yy) = (G(Xx, Yy) * (1 - AB)) + (Lg * AB)
B(Xx, Yy) = (B(Xx, Yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, Xx, k + Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub SinusLineY(Rx1%, Ry1%, Rx2%, Ry2%, AB!, Wave%, Ampl%, LCol&, Dist%, Eff%)
On Error Resume Next
Dim Degree As Single, k!
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
For Yy = Ry1 To Ry2 - 1
For Xx = Rx1 To Rx2 - 1 Step Dist
Degree = Yy * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
Color = GetPixel(FMain.Pic1.hdc, k + Xx, Yy)
r(Xx, Yy) = Color Mod 256&
G(Xx, Yy) = ((Color And &HFF00) / 256&) Mod 256&
B(Xx, Yy) = (Color And &HFF0000) / 65536
r(Xx, Yy) = (r(Xx, Yy) * (1 - AB)) + (Lr * AB)
G(Xx, Yy) = (G(Xx, Yy) * (1 - AB)) + (Lg * AB)
B(Xx, Yy) = (B(Xx, Yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, k + Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Xx
FMain.PB1.Value = Xx - Rx1
Next Yy
EndProcess
End Sub

Public Sub SBorder(Dist%, AB!, LCol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Line (Xx, Xx)-(FMain.TempMem.Width - 1 - Xx, FMain.TempMem.Height - 1 - Xx), LCol, B
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub GBorder1(Dist%, AB!, LCol&, Scol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
BI = (Gb1 - Gb) / Dist
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Line (Xx, Xx)-(FMain.TempMem.Width - 1 - Xx, FMain.TempMem.Height - 1 - Xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub GBorder2(Dist%, AB!, LCol&, Scol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
BI = (Gb1 - Gb) / Dist * 2
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist / 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For Xx = 0 To (Dist / 2)
FMain.TempMem.Line (Xx, Xx)-(FMain.TempMem.Width - 1 - Xx, FMain.TempMem.Height - 1 - Xx), RGB(Gr, Gg, Gb), B
FMain.TempMem.Line (Dist - Xx, Dist - Xx)-(FMain.TempMem.Width - 1 - Dist + Xx, FMain.TempMem.Height - 1 - Dist + Xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub CBorder(Dist%, AB!, LCol&)
On Error Resume Next
Dim Ra!, Rr%
If LCol = 0 Then LCol = &H10101
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - Xx, LCol, , , Ra
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCBorder1(Dist%, AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI
Dim Ra!, Rr%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
BI = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - Xx, RGB(Gr, Gg, Gb), , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCBorder2(Dist%, AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI
Dim Ra!, Rr%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
BI = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For Xx = 0 To Dist / 2
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - Xx, RGB(Gr, Gg, Gb), , , Ra
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - Dist + Xx, RGB(Gr, Gg, Gb), , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub MixSolid(AB!, LCol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = LCol
BeginProcess
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixGradient1(AB!, LCol&, Scol&, ch%)
On Error Resume Next
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, BI, Dist%
If ch = 0 Then Dist = FMain.Pic1.Height
If ch = 1 Then Dist = FMain.Pic1.Width
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
BI = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
BeginProcess
For Xx = 0 To Dist - 1
If ch = 0 Then FMain.TempMem.Line (0, Xx)-(FMain.TempMem.Width - 1, Xx), RGB(Gr, Gg, Gb)
If ch = 1 Then FMain.TempMem.Line (Xx, 0)-(Xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixGradient2(AB!, LCol&, Scol&, ch%)
On Error Resume Next
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, BI, Dist%
If ch = 0 Then Dist = FMain.Pic1.Height
If ch = 1 Then Dist = FMain.Pic1.Width
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
BI = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
BeginProcess
For Xx = 0 To Dist / 2
If ch = 0 Then
FMain.TempMem.Line (0, Xx)-(FMain.TempMem.Width - 1, Xx), RGB(Gr, Gg, Gb)
FMain.TempMem.Line (0, Dist - 1 - Xx)-(FMain.TempMem.Width - 1, Dist - 1 - Xx), RGB(Gr, Gg, Gb)
Else
FMain.TempMem.Line (Xx, 0)-(Xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
FMain.TempMem.Line (Dist - 1 - Xx, 0)-(Dist - 1 - Xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
End If
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixBoxGradient1(AB!, LCol&, Scol&)
On Error Resume Next
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, BI, Dist%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If FMain.Pic1.Width < FMain.Pic1.Height Then
Dist = FMain.Pic1.Width / 2
Else
Dist = FMain.Pic1.Height / 2
End If
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
BI = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Line (Xx, Xx)-(FMain.TempMem.Width - 1 - Xx, FMain.TempMem.Height - 1 - Xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixBoxGradient2(AB!, LCol&, Scol&)
On Error Resume Next
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, BI, Dist%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If FMain.Pic1.Width < FMain.Pic1.Height Then
Dist = FMain.Pic1.Width / 2
Else
Dist = FMain.Pic1.Height / 2
End If
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
BI = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
BeginProcess
For Xx = 0 To Dist / 2
FMain.TempMem.Line (Xx, Xx)-(FMain.TempMem.Width - 1 - Xx, FMain.TempMem.Height - 1 - Xx), RGB(Gr, Gg, Gb), B
FMain.TempMem.Line (Dist - Xx, Dist - Xx)-(FMain.TempMem.Width - 1 - Dist + Xx, FMain.TempMem.Height - 1 - Dist + Xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub GCircle1(AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI, Dist%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
    Dist = Sqr(((FMain.Pic1.Width / 2) ^ 2) + ((FMain.Pic1.Height / 2) ^ 2))
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
BI = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
BeginProcess
For Xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Dist - Xx, RGB(Gr, Gg, Gb) ', , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCircle2(AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
'FIXIT: Declare 'Ri' and 'Gi' and 'Bi' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, BI, Dist%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
    Dist = Sqr(((FMain.Pic1.Width / 2) ^ 2) + ((FMain.Pic1.Height / 2) ^ 2))
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
BI = (Gb1 - Gb) / Dist * 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
BeginProcess
FMain.PB1.Max = Dist / 2
For Xx = 0 To Dist / 2
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Dist - Xx, RGB(Gr, Gg, Gb) ', , , Ra
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Xx, RGB(Gr, Gg, Gb) ', , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + BI
FMain.PB1.Value = Xx
Next Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub MixPic(AB!, Op As Boolean)
On Error Resume Next
BeginProcess
Dim LCol&
LCol = 0
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.BackColor = 0
If Op = True Then
FMain.TempMem.PaintPicture FPicture.Pic2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Else
FMain.TempMem.PaintPicture FPicture.Pic2, (FMain.Pic1.Width - FPicture.Pic2.Width) / 2, (FMain.Pic1.Height - FPicture.Pic2.Height) / 2, FPicture.Pic2.Width, FPicture.Pic2.Height
End If
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixPattern(AB!)
BeginProcess
Dim LCol&
LCol = 0
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.BackColor = 0
For Xx = 0 To FMain.Pic1.Width / FPicture.Pic2.Width
For Yy = 0 To FMain.Pic1.Height / FPicture.Pic2.Height
FMain.TempMem.PaintPicture FPicture.Pic2, Xx * FPicture.Pic2.Width, Yy * FPicture.Pic2.Height, FPicture.Pic2.Width, FPicture.Pic2.Height
Next Yy, Xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub Echo(ENr%, ERed%, ex%, EY%) 'echo picture
Dim EchoW&, EchoH&
Dim EchoLeft%, EchoTop%, Phase%
On Error Resume Next
FMain.Label2.Caption = "": DoEvents
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.Picture = FMain.Pic1.Image
EchoW = FMain.TempMem.Width - 1
EchoH = FMain.TempMem.Height - 1
Phase = 0
FMain.PB1.Max = ENr - 1
For Xx = 0 To ENr - 1
If FEcho.Check1.Value = 1 Then Phase = Xx
EchoW = EchoW * (100 - ERed) / 100
EchoH = EchoH * (100 - ERed) / 100
EchoLeft = (FMain.TempMem.Width / 2) - (EchoW / 2) + ((Phase + 1) * ex)
EchoTop = (FMain.TempMem.Height / 2) - (EchoH / 2) + ((Phase + 1) * EY)
FMain.Pic1.PaintPicture FMain.TempMem, EchoLeft, EchoTop, EchoW, EchoH
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Mozaic(Rx1, Ry1, Rx2, Ry2, Br%) 'mozaic
Dim Br2%, MozaicColor&, qq%, pp%
On Error Resume Next
Br2 = Int(Br / 2)
FMain.PB1.Max = Rx2 - Rx1
BeginProcess
For Xx = Rx1 To Rx2 Step Br
For Yy = Ry1 To Ry2 Step Br
MozaicColor = GetPixel(FMain.Pic1.hdc, Xx + Br2, Yy + Br2)
    For qq = Xx To Xx + Br - 1
    For pp = Yy To Yy + Br - 1
    r(qq, pp) = MozaicColor
    Next pp, qq
Next Yy, Xx
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, Xx, Yy, r(Xx, Yy)
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

'FIXIT: Declare 'Rx1' and 'Ry1' and 'Rx2' and 'Ry2' with an early-bound data type          FixIT90210ae-R1672-R1B8ZE
Public Sub Mozaic2(Rx1, Ry1, Rx2, Ry2, Br%) 'mozaic2
Dim Br2%, MozaicColor&, qq%, pp%, R1&, G1&, B1&
On Error Resume Next
Br2 = Int(Br / 2)
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For Xx = Rx1 To Rx2 Step Br
For Yy = Ry1 To Ry2 Step Br
Color = GetPixel(FMain.Pic1.hdc, Xx + Br2, Yy + Br2)
R1 = Color Mod 256&
G1 = ((Color And &HFF00) / 256&) Mod 256&
B1 = (Color And &HFF0000) / 65536
    For qq = Xx To Xx + Br - 1
    For pp = Yy To Yy + Br - 1
    If qq = Xx Or pp = Yy Or qq = Xx + Br - 1 Or pp = Yy + Br - 1 Then
        r(qq, pp) = r(qq, pp) - ((Rnd * 10) - 5)
        If r(qq, pp) < 0 Then r(qq, pp) = 0
        G(qq, pp) = G(qq, pp) - ((Rnd * 10) - 5)
        If G(qq, pp) < 0 Then G(qq, pp) = 0
        B(qq, pp) = B(qq, pp) - ((Rnd * 10) - 5)
        If B(qq, pp) < 0 Then B(qq, pp) = 0
    Else
    r(qq, pp) = R1
    G(qq, pp) = G1
    B(qq, pp) = B1
    End If
    Next pp, qq
Next Yy, Xx
For Xx = Rx1 To Rx2
For Yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy), G(Xx, Yy), B(Xx, Yy))
Next Yy
FMain.PB1.Value = Xx - Rx1
Next Xx
EndProcess
End Sub

Public Sub EffectX(Strength%, Wave As Single, Eff%) 'wave x
On Error Resume Next
Dim Degree As Single, k!
ReDim LineColor(FMain.Pic1.Width)
BeginProcess
Wave = Wave / 10
FMain.PB1.Max = FMain.Pic1.Height - 1
For Yy = 0 To FMain.Pic1.Height - 1
Degree = (Yy * Wave / 180 * 3.14)
If Eff = 0 Then k = (Cos(Degree) * Strength)
If Eff = 1 Then k = Abs(Cos(Degree) * Strength)
GetColorsX
For Xx = 0 To FMain.Pic1.ScaleWidth
If k < 0 Then
SetPixel FMain.Pic1.hdc, Xx, Yy, LineColor(FMain.Pic1.ScaleWidth - 1)
Else
SetPixel FMain.Pic1.hdc, Xx, Yy, LineColor(0)
End If
Next Xx
For Xx = 0 To FMain.Pic1.ScaleWidth - 1
SetPixel FMain.Pic1.hdc, Xx + k, Yy, LineColor(Xx)
Next Xx
FMain.PB1.Value = Yy
Next Yy
EndProcess
End Sub

Public Sub EffectY(Strength%, Wave As Single, Eff%) 'wave y
On Error Resume Next
Dim k!
ReDim LineColor(FMain.Pic1.Height)
BeginProcess
Wave = Wave / 10
FMain.PB1.Max = FMain.Pic1.Width - 1
For Xx = 0 To FMain.Pic1.Width - 1
If Eff = 0 Then k = (Cos(Xx * Wave / 180 * 3.14) * Strength)
If Eff = 1 Then k = Abs(Cos(Xx * Wave / 180 * 3.14) * Strength)
GetColorsY
For Yy = 0 To FMain.Pic1.ScaleHeight
If k < 0 Then
SetPixel FMain.Pic1.hdc, Xx, Yy, LineColor(OB.ScaleHeight - 1)
Else
SetPixel FMain.Pic1.hdc, Xx, Yy, LineColor(0)
End If
Next Yy
For Yy = 0 To FMain.Pic1.ScaleHeight - 1
SetPixel FMain.Pic1.hdc, Xx, Yy + k, LineColor(Yy)
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

Private Sub GetColorsX()
Dim tt%
For tt = 0 To FMain.Pic1.Width - 1
LineColor(tt) = GetPixel(FMain.Pic1.hdc, tt, Yy)
Next tt
End Sub

Private Sub GetColorsY()
Dim tt%
For tt = 0 To FMain.Pic1.Height - 1
LineColor(tt) = GetPixel(FMain.Pic1.hdc, Xx, tt)
Next tt
End Sub

Public Sub KillColXGrad1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border left 1
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width / 4
For Xx = Rx1 To Rx2 / 4
kk = (1 - (Xx / FMain.Pic1.ScaleWidth) * 4)
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

Public Sub KillColXGrad2(Rx1%, Ry1%, Rx2%, Ry2%)  'grad border left 2
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width / 2
For Xx = Rx1 To Rx2 / 2
kk = (1 - (Xx / FMain.Pic1.ScaleWidth) * 2)
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

Public Sub KillColXGrad3(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border left 3
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width - 1
For Xx = Rx1 To Rx2 - 1
kk = 1 - (Xx / FMain.Pic1.ScaleWidth)
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

Public Sub KillColXGradRev1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 1
On Error Resume Next
FMain.PB1.Max = Rx2 - 1 - (Rx2 / 4 * 3)
For Xx = Rx2 / 4 * 3 To Rx2 - 1
kk = (Xx - (Rx2 / 4 * 3)) / (Rx2 / 4)
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx - (Rx2 / 4 * 3)
Next Xx
EndProcess
End Sub

Public Sub KillColXGradRev2(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 2
On Error Resume Next
FMain.PB1.Max = Rx2 - 1 - (Rx2 / 2)
For Xx = Rx2 / 2 To Rx2 - 1
kk = (Xx - (Rx2 / 2)) / (Rx2 / 2)
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx - (Rx2 / 2)
Next Xx
EndProcess
End Sub

Public Sub KillColXGradRev3(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 3
On Error Resume Next
FMain.PB1.Max = Rx2
For Xx = Rx1 To Rx2 - 1
kk = Xx / Rx2
For Yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub

Public Sub KillColYGrad1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border top 1
On Error Resume Next
FMain.PB1.Max = Ry2 / 4
For Yy = Ry1 To Ry2 / 4
kk = (1 - (Yy / Ry2) * 4) '/ 1.4
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy
Next Yy
EndProcess
End Sub

Public Sub KillColYGrad2(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border top 2
On Error Resume Next
FMain.PB1.Max = Ry2 / 2
For Yy = Ry1 To Ry2 / 2
kk = (1 - (Yy / Ry2) * 2)
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy
Next Yy
EndProcess
End Sub

Public Sub KillColYGrad3(Rx1%, Ry1%, Rx2%, Ry2%)  'grad top border 3
On Error Resume Next
FMain.PB1.Max = Ry2
For Yy = Ry1 To Ry2 - 1
kk = 1 - (Yy / Ry2)
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy
Next Yy
EndProcess
End Sub

Public Sub KillColYGradRev1(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 1
On Error Resume Next
FMain.PB1.Max = Ry2 - 1 - (Ry2 / 4 * 3)
For Yy = ((Ry2) / 4) * 3 To Ry2 - 1
kk = (Yy - (Ry2 / 4 * 3)) / (Ry2 / 4)
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy - (Ry2 / 4 * 3)
Next Yy
EndProcess
End Sub

Public Sub KillColYGradRev2(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 2
On Error Resume Next
FMain.PB1.Max = Ry2 - 1 - (Ry2 / 2)
For Yy = ((Ry2) / 2) To Ry2 - 1
kk = (Yy - (Ry2 / 2)) / (Ry2 / 2)
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy - (Ry2 / 2)
Next Yy
EndProcess
End Sub

Public Sub KillColYGradRev3(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 3
On Error Resume Next
FMain.PB1.Max = Ry2 - 1
For Yy = Ry1 To Ry2 - 1
kk = Yy / Ry2
For Xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, Xx, Yy, RGB(r(Xx, Yy) - (r(Xx, Yy) * kk), G(Xx, Yy) - (G(Xx, Yy) * kk), B(Xx, Yy) - (B(Xx, Yy) * kk))
Next Xx
FMain.PB1.Value = Yy
Next Yy
EndProcess
End Sub

Public Sub Tile(XTile%, YTile%)
On Error Resume Next
Dim TileX%, TileY%
FMain.PB1.Max = XTile - 1
TileX = Int(FMain.Pic1.Width / XTile)
TileY = Int(FMain.Pic1.Height / YTile)
FMain.Label2.Caption = "": DoEvents
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.Picture = FMain.Pic1.Image
For Xx = 0 To XTile - 1
For Yy = 0 To YTile - 1
FMain.Pic1.PaintPicture FMain.TempMem, Xx * TileX, Yy * TileY, TileX, TileY
Next Yy
FMain.PB1.Value = Xx
Next Xx
EndProcess
End Sub
