VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl Command 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   ToolboxBitmap   =   "XpButton.ctx":0000
   Begin VB.PictureBox imgButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   20
      Left            =   1440
      Top             =   3120
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":0312
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   1
      Left            =   0
      Top             =   960
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":19B4
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   2
      Left            =   0
      Top             =   1320
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":3056
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   17
      Left            =   1440
      Top             =   2040
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":46F8
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   8
      Left            =   0
      Top             =   3480
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":5D9A
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   25
      Left            =   1440
      Top             =   4920
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":743C
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   22
      Left            =   1440
      Top             =   3840
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":8ADE
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   11
      Left            =   0
      Top             =   4560
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":A180
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   10
      Left            =   0
      Top             =   4200
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":B822
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   14
      Left            =   1440
      Top             =   960
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":CEC4
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   21
      Left            =   1440
      Top             =   3480
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":E566
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   23
      Left            =   1440
      Top             =   4200
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":FC08
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   9
      Left            =   0
      Top             =   3840
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":112AA
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   13
      Left            =   1440
      Top             =   600
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1294C
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   5
      Left            =   0
      Top             =   2400
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":13FEE
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   19
      Left            =   1440
      Top             =   2760
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":15690
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   12
      Left            =   0
      Top             =   4920
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":16D32
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   7
      Left            =   0
      Top             =   3120
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":183D4
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   4
      Left            =   0
      Top             =   2040
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":19A76
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   15
      Left            =   1440
      Top             =   1320
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1B118
   End
   Begin PicClip.PictureClip pc 
      Left            =   0
      Top             =   240
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1C7BA
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   18
      Left            =   1440
      Top             =   2400
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1D398
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   16
      Left            =   1440
      Top             =   1680
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1DF76
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   0
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   3
      Left            =   0
      Top             =   1680
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":1F618
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   6
      Left            =   0
      Top             =   2760
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":20CBA
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   24
      Left            =   1440
      Top             =   4560
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":2235C
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   0
      Left            =   0
      Top             =   600
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "XpButton.ctx":239FE
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Xp button"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   705
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINT_API) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Declare Function DrawFocusRect Lib "User32" (ByVal hdc As Long, lpRect As GIVEFOCUS) As Long
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*   *  This code is originaly writen by Aleks Kimhi - Aki.         *   *   *                             *
'  *    As you might have read that this ActiveX control is free,   *   *         O                       *
'*   *  and therefore can be distributed for free as long as you      *                                       *
'  *    do not sell it for profit and there's still my name on it.         *                     ____|            *
'*   *  I've spend some time on this project so please                *  *                                        *
'  *    don't just take it, use it and say you wrote it.                 *  *   *                                     *
'*   *  Thank you for your co-operation. Any comments,        *   *  *   *      P                          *
'  *    good or bad, would be greatly appreciated.                       *                                      *
'*   *  E -mail: aniram@ zahav.net.il                                       *    *                                 *
'  *    Tel-Aviv, Israel    A & M © Copyright 2002                   *    *     *                           *
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
' *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type GIVEFOCUS
  Left1 As Long
  Top1 As Long
  Right1 As Long
  Bottom1 As Long
End Type

Public Enum ButtonPicture
    DefaultButton = 0: XP_Black = 1: XP_Blood = 2: XP_BlueHighlight = 3: XP_BlueMetalic = 4
    XP_Darker = 5: XP_Disco = 6: XP_Gold = 7: XP_Grass = 8: XP_Hot = 9: XP_Lady = 10
    XP_LightGreen = 11: XP_Lily = 12: XP_Limon = 13: XP_Hawai = 14: XP_Ocean = 15
    XP_OldStyle = 16: XP_Orange = 17: XP_Original = 18: XP_Paper = 19: XP_Picaso = 20
    XP_Rain = 21: XP_Red = 22: XP_Silver = 23: XP_Wood = 24: XP_Yellow = 25
End Enum

Public Enum Alignment
    None1 = 0
    Left1 = 1
    Right1 = 2
End Enum

Dim mPic As ButtonPicture
Const defPic = ButtonPicture.DefaultButton
Dim btnDown, gotFocus, a As Integer
Dim mFont As Font
Dim mForeColor As OLE_COLOR
Const defForeColor = vbBlack
Dim Ftime, mEnabled As Boolean 'Ftime is giving focus to button(only to one-on the beggining)
Dim Iam As Byte '(0 Normal)-(1 Clicked)-(2 Disabled)-(3 On the button)-(4 Last clicked)
Dim mPicAlign As Alignment
Const defPicAlign = None1

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then Exit Sub
        If Iam = 3 Then Exit Sub
            If btnDown = 1 Then Exit Sub
        Iam = 3: drawPic (3): Timer1.Enabled = True
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
        If mEnabled = False Then Exit Sub
            If btnDown = 1 Then Exit Sub
                lbl.Top = lbl.Top + 1.01: lbl.Left = lbl.Left + 1.01
            imgButton.Top = imgButton.Top + 1.01: imgButton.Left = imgButton.Left + 1.01
        btnDown = 1: Iam = 1: drawPic (1)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then Exit Sub
        SetLabel
            SetAlign
                Iam = 3: drawPic (3): btnDown = 0
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If btnDown = 1 Then Exit Sub
    RaiseEvent KeyPress(KeyAscii)
    RaiseEvent Click
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    Call UserControl_MouseUp(0, 0, 0, 0)
    Call UserControl_KeyPress(KeyCode)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
            RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Private Sub imgButton_Click()
    UserControl_Click
End Sub

Private Sub imgButton_DblClick()
    RaiseEvent DblClick
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Private Sub lbl_Click()
    UserControl_Click
End Sub

Private Sub lbl_DblClick()
    RaiseEvent DblClick
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
    ButtonLook = defPic: Iam = 0
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True: ButtonLook = defPic: Set Font = UserControl.Ambient.Font
    ForeColor = defForeColor: Set Icon = imgButton.Picture
End Sub

Private Sub UserControl_Resize()
    CheckEnabled
End Sub

Private Sub CheckEnabled()
    If mEnabled = False Then
        Iam = 2: drawPic (2): lbl.Enabled = False: imgButton.Visible = False
            mPicAlign = None1: SetLabel
                Exit Sub
        Else
          Iam = 0: drawPic (0): lbl.Enabled = True
                If mPicAlign = None1 Then
                    imgButton.Visible = False
                        Else
                            imgButton.Visible = True
                End If
    End If
    SetLabel
    SetAlign
End Sub

Private Sub drawPic(z As Integer) 'Draw button
   
    UserControl.ScaleMode = 3 'Draw in pixels
    Dim Xx, Yy, W, H As Integer
                                                                                                     
    Xx = UserControl.ScaleWidth - 3
    Yy = UserControl.ScaleHeight - 3
    W = UserControl.ScaleWidth - 6
    H = UserControl.ScaleHeight - 6
       
    UserControl.PaintPicture pc.GraphicCell(z), 0, 0, 3, 3, 0, 0, 3, 3 'left top corner
    UserControl.PaintPicture pc.GraphicCell(z), Xx, 0, 3, 3, 15, 0, 3, 3 'right top corner
    UserControl.PaintPicture pc.GraphicCell(z), Xx, Yy, 3, 3, 15, 18, 3, 3 'right down corner
    UserControl.PaintPicture pc.GraphicCell(z), 0, Yy, 3, 3, 0, 18, 3, 3 'left down corner
    UserControl.PaintPicture pc.GraphicCell(z), 3, 0, W, 3, 3, 0, 12, 3 'top line
    UserControl.PaintPicture pc.GraphicCell(z), Xx, 3, 3, H, 15, 3, 3, 15 'right line
    UserControl.PaintPicture pc.GraphicCell(z), 3, Yy, W, 3, 3, 18, 12, 3 'bottom line
    UserControl.PaintPicture pc.GraphicCell(z), 0, 3, 3, H, 0, 3, 3, 15 'left line
    UserControl.PaintPicture pc.GraphicCell(z), 3, 3, W, H, 3, 3, 12, 15 'and fill
    
'   If Ftime = True Then
            If Iam = 1 And btnDown = 1 Or _
            Iam = 3 And gotFocus = 1 Or _
            Iam = 4 And gotFocus = 1 Then DrawFocus
'            End If
    End Sub

Private Sub MakeMeHappy(ch As Integer)
    pc.Picture = pcChoice(ch).Picture
End Sub

Private Sub imgButton_GotFocus()
    Call UserControl_GotFocus
End Sub

Private Sub imgButton_LostFocus()
    Call UserControl_LostFocus
End Sub

Private Sub UserControl_GotFocus()
    If mEnabled = False Then Exit Sub
        gotFocus = 1
    If btnDown = 0 Then: Iam = 4: drawPic (4): Else Iam = 4
End Sub
                                                         
Private Sub UserControl_LostFocus()
    If mEnabled = False Then Exit Sub
        gotFocus = 0: Iam = 0: drawPic (0)
End Sub

Private Sub SetLabel()
    If mPicAlign = None1 Then
        lbl.Left = ((UserControl.ScaleWidth) - lbl.Width) / 2
            ElseIf mPicAlign = Left1 Then
                lbl.Left = imgButton.Width + 7
            ElseIf mPicAlign = Right1 Then
        lbl.Left = 5
   End If
   lbl.Top = ((UserControl.ScaleHeight) - lbl.Height) / 2
End Sub

Private Sub SetAlign()
    If mPicAlign = Left1 Then
        imgButton.Left = 5
            Else
                imgButton.Left = (UserControl.ScaleWidth - imgButton.Width) - 5
    End If
    imgButton.Top = ((UserControl.ScaleHeight) - imgButton.Height) / 2
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             'Aki
    mPicAlign = PropBag.ReadProperty("IconAlign", defPicAlign)
    Set Icon = PropBag.ReadProperty("Icon", imgButton.Picture)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption", "XpButton")
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    FontColor = PropBag.ReadProperty("FontColor", defForeColor)
    ButtonLook = PropBag.ReadProperty("ButtonLook", defPic)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("IconAlign", mPicAlign, defPicAlign)
    Call PropBag.WriteProperty("Icon", imgButton.Picture, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "XpButton1")
    Call PropBag.WriteProperty("Font", mFont, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("FontColor", mForeColor, defForeColor)
    Call PropBag.WriteProperty("ButtonLook", mPic, defPic)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    UserControl.Enabled() = NewEnabled: mEnabled = NewEnabled: CheckEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
    Set Font = mFont
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set mFont = NewFont: Set UserControl.Font = NewFont: Set lbl.Font = mFont
    PropertyChanged "Font"
    SetLabel
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = mForeColor
End Property

Public Property Let FontColor(ByVal NewFontColor As OLE_COLOR)
    mForeColor = NewFontColor: lbl.ForeColor = mForeColor
    PropertyChanged "FontColor"
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lbl.Caption() = NewCaption
    PropertyChanged "Caption"
End Property

Public Property Get ButtonLook() As ButtonPicture
    ButtonLook = mPic
End Property

Public Property Let ButtonLook(ByVal NewButtonLook As ButtonPicture)
    mPic = NewButtonLook: MakeMeHappy (mPic): CheckEnabled
    PropertyChanged "ButtonLook"
End Property

Public Property Get IconAlign() As Alignment
    IconAlign = mPicAlign
End Property

Public Property Let IconAlign(ByVal NewPicAlign As Alignment)
    If NewPicAlign = None1 Then imgButton.Visible = False: mPicAlign = 0: SetLabel: Exit Property
    mPicAlign = NewPicAlign
    PropertyChanged "IconAlign"
    SetLabel
    SetAlign
End Property

                                                         'You can use or picture or image.
  Public Property Get Icon() As Picture 'If you have circle and you don't wan't to see
       Set Icon = imgButton.Picture        'backcolor then use image( but your picture will
  End Property                                    'blink on mouse click). From that reason,
                                                         'picture is the answer.

Public Property Set Icon(ByVal NewIcon As Picture)
    Set imgButton.Picture = LoadPicture()
    Set imgButton.Picture = NewIcon
    If NewIcon = 0 Then Exit Property
    PropertyChanged "Icon"
    If mPicAlign = None1 Then: mPicAlign = Left1
    SetLabel
    SetAlign
    imgButton.Visible = True
End Property

Private Sub Timer1_Timer()
    Dim dot As POINT_API
    UserControl.ScaleMode = 3 'must have this 'cause of x and y, to know how to calc
    Call GetCursorPos(dot) 'get mouse position
        ScreenToClient UserControl.hWnd, dot 'must have
  
            'checking if mouse is on our control, by x and y
            If dot.X < UserControl.ScaleLeft Or _
                dot.Y < UserControl.ScaleTop Or _
                    dot.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
                        dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
                        
                            Timer1.Enabled = False
                                    If gotFocus = 1 Then
                                        Iam = 4: drawPic (4)
                                            Else: Call UserControl_LostFocus
                                    End If
                        RaiseEvent MouseOut
            End If
End Sub

Private Sub DrawFocus()
  Dim F As GIVEFOCUS
  
        F.Top1 = 2
            F.Right1 = 2 + UserControl.ScaleWidth - 4
                F.Bottom1 = UserControl.ScaleHeight - 2
                    F.Left1 = 2
                    
        DrawFocusRect UserControl.hdc, F
End Sub
