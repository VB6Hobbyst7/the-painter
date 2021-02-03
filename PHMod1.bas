Attribute VB_Name = "PHMod1"
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'FIXIT: Declare 'r' and 'G' and 'B' with an early-bound data type                          FixIT90210ae-R1672-R1B8ZE
Public r(), G(), B(), Tim%, PicFileName$, Temp$, FTitle$
Public Xx%, Yy%, Xcor0%, Xcor1%, Ycor0%, Ycor1%, XXX1%, XXX2%, YYY1%, YYY2%
Public Col%, Mix%
Public PicMem As Picture, Im As Picture
Public PicMem0(4) As Picture, MemCount%
Public OrWidth%, OrHeight%, Factor!
'FIXIT: Declare 'Scol' with an early-bound data type                                       FixIT90210ae-R1672-R1B8ZE
Public Scol(15)
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const LF_FACESIZE = 32
Public Const LOGPIXELSY = 90
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lsngStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lsngPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type


Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As LOGFONT, FontName As String, ZeroPos As Long
    CopyMemory LF, ByVal lplf, LenB(LF)
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, Chr$(0))
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    FText.Combo1.AddItem FontName
    EnumFontProc = 1
End Function
   
'FIXIT: Declare 'T3D' and 'Obj0' and 'Obj' with an early-bound data type                   FixIT90210ae-R1672-R1B8ZE
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim r%, G%, B%, R1%, G1%, B1%, R2%, G2%, b2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
r = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = r + 64
    If R1 > 255 Then R1 = 255
    R2 = r - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    b2 = B - 64
    If b2 < 0 Then b2 = 0
    B3 = B1
    B4 = b2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = r + 64
    If R1 > 255 Then R1 = 255
    R2 = r - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    b2 = B - 64
    If b2 < 0 Then b2 = 0
    B4 = B1
    B3 = b2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = r + 64
    If R2 > 255 Then R2 = 255
    R1 = r - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    b2 = B + 64
    If b2 > 255 Then b2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = b2
End If
If Style3D = 3 Then 'InsetInset
    R2 = r + 64
    If R2 > 255 Then R2 = 255
    R1 = r - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    b2 = B + 64
    If b2 > 255 Then b2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = b2
End If
If Style3D = 4 Then 'No Border
R1 = r: R2 = r: R3 = r: R4 = r
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: b2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, b2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, b2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, b2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, b2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

Public Sub SetScrollBars()
With FMain
.Pic1.Move 0, 0
.VS1.Value = 0
.HS1.Value = 0
.VS1.Enabled = False
.HS1.Enabled = False
If .Pic1.Width > .PicX.Width - .VS1.Width Then .HS1.Enabled = True
If .Pic1.Height > .PicX.Height - .HS1.Height Then .VS1.Enabled = True
    If .VS1.Enabled = True Then
    .VS1.Max = .PicX.Height - .HS1.Height - .Pic1.Height
    .VS1.LargeChange = .Pic1.Height / 10
    End If
    If .HS1.Enabled = True Then
    .HS1.Max = .PicX.Width - .VS1.Width - .Pic1.Width
    .HS1.LargeChange = .Pic1.Width / 10
    End If
End With
End Sub

Public Sub SetPicInfo()
With FMain
    .Label1.Caption = vbCr ' " Í¼ÏóÐÅÏ¢£º" & vbCr &
    If PicFileName = "" Then
    .Label1.Caption = .Label1.Caption & vbCr
    Else
    .Label1.Caption = .Label1.Caption & lgT(305) & PicFileName & vbCr
    End If
    .Label1.Caption = .Label1.Caption & lgT(306) & FMain.Pic1.Width & vbCr
    .Label1.Caption = .Label1.Caption & lgT(307) & FMain.Pic1.Height & vbCr


'.Caption = PicFileName & " - " & lgT(314)
ReDim r(.Pic1.Width, .Pic1.Height)
ReDim G(.Pic1.Width, .Pic1.Height)
ReDim B(.Pic1.Width, .Pic1.Height)
OrHeight = .Pic1.Height
OrWidth = .Pic1.Width
If .Pic1.Width > 900 And .Pic1.Height > 800 Then MsgBox lgT(308), vbInformation, "Too Large"
End With
End Sub

Public Sub SaveRedo()
For Xx = 4 To 1 Step -1
FMain.TempMem = PicMem0(Xx - 1)
Set PicMem0(Xx) = FMain.TempMem.Image
Next Xx
FMain.TempMem.Picture = FMain.Pic1.Image
Set PicMem0(0) = FMain.TempMem.Image
MemCount = MemCount + 1
If MemCount > 5 Then MemCount = 5
ShowMem
End Sub

Public Sub Redo()
FMain.Pic1 = PicMem0(0)
For Xx = 0 To 3
Set PicMem0(Xx) = PicMem0(Xx + 1)
Next Xx
Set PicMem0(4) = Nothing
MemCount = MemCount - 1
If MemCount = 0 Then MemCount = 0
ShowMem
SetPicInfo
End Sub
Public Sub ShowMem()


'On Error Resume Next
FMain.Toolbar1.Buttons(1).Enabled = False
FMain.Cancel1.Enabled = False
',,,,,,,,,,,,,,
For i = 0 To 4
ToolRedo.Image1(i).Enabled = False
FMain.Image1(i).Enabled = False
Next
''''''''''''
For Xx = 0 To 4
ToolRedo.Image1(Xx) = PicMem0(Xx)
FMain.Image1(Xx) = PicMem0(Xx)

Next Xx






If MemCount > 0 Then
FMain.Toolbar1.Buttons(1).Enabled = True
FMain.Cancel1.Enabled = True
',,,,,,,,,,,,
Dim Doe As Integer
Doe = MemCount - 1
For s = 0 To Doe
ToolRedo.Image1(s).Enabled = True
FMain.Image1(s).Enabled = True
Next
''''
End If
End Sub

Public Sub ClearMem()
For Xx = 0 To 4
Set PicMem0(Xx) = Nothing
Next Xx
ShowMem
End Sub

Public Sub SetCoordinates()
ToolXY.Label4.Caption = lgT(309) & vbCr & Format(Xcor0, "000") & " X " & Format(Ycor0, "000") & "         " & Format(Xcor1, "000") & " X " & Format(Ycor1, "000")
FMain.Label4.Caption = ToolXY.Label4.Caption
Xcor1 = Xcor0 + Xcor1
Ycor1 = Ycor0 + Ycor1
End Sub

