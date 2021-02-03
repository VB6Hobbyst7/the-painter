VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T"
   ClientHeight    =   3660
   ClientLeft      =   4140
   ClientTop       =   3450
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000B&
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   3555
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关于"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   0
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced Creating"
      Height          =   1215
      Left            =   240
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   1080
         TabIndex        =   25
         Text            =   "#"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   2280
         MousePointer    =   3  'I-Beam
         TabIndex        =   22
         Text            =   "255"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   1680
         MousePointer    =   3  'I-Beam
         TabIndex        =   21
         Text            =   "192"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   1080
         MousePointer    =   3  'I-Beam
         TabIndex        =   20
         Text            =   "170"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Hex:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "RGB:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Display color"
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   3135
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   13
         ToolTipText     =   "RGB"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "16进制颜色值"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DSP"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "RGB"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hex"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create a color"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1320
         Max             =   255
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   240
         Value           =   128
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   1320
         Max             =   255
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   600
         Value           =   128
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   1320
         Max             =   255
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   960
         Value           =   128
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   720
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   720
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   720
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Width           =   600
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   975
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      MousePointer    =   2  'Cross
      Picture         =   "frmColor.frx":0CCE
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   1
      ToolTipText     =   "Drag to the color place"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Instruction"
      Height          =   1095
      Left            =   1200
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Declare Function SetWindowPos Lib "User32" ( _
    ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
  Dim z As POINTAPI
  Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

 Dim a, B, C, d, e, F, G, H, isf As Long   '声明长整形变量,用于存储临时色值
 

Private Sub Check1_Click()
If Check1.Value = 0 Then
Frame1.Visible = False
ElseIf Check1.Value = 1 Then
Frame1.Visible = True
End If
End Sub

Private Sub Command1_Click()
Me.Hide
about1.Show 1
Me.Show
End Sub

Private Sub Command2_Click()
ShellExecute Me.hWnd, "Open", App.Path + "\Help\color.htm", "", App.Path, 1
End Sub

Private Sub Form_Initialize()
Me.Caption = "取色吸管 - 小画家"
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE _
Or SWP_NOSIZE
 ' Picture1.MouseIcon = LoadResPicture(104, vbResIcon)
 ' frmMain.WindowState = 1
 If LangA = "lge" Then
 Label2.Caption = "Instruction:" & vbCrLf & "Mouse down the image left, and drag to the color you need, then release mouse."
 Else
 Label2.Caption = "说明：" & vbCrLf & "按住左边的图片，然后拖动到你希望取得颜色的点，再放开鼠标。"
 Frame3.Caption = "颜色显示"
 Frame2.Caption = "创建颜色"
  End If
  
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub



Private Sub Frame1_Click()
   Frame1.Visible = False
End Sub




Private Sub HScroll1_Change()
  If isf <> 1 Then
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text2.Text = HScroll1.Value
  End If
End Sub

Private Sub HScroll1_GotFocus()
  isf = 0
End Sub

Private Sub HScroll2_Change()
  If isf <> 1 Then
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text3.Text = HScroll2.Value
  End If
End Sub

Private Sub HScroll2_GotFocus()
  isf = 0
End Sub

Private Sub HScroll3_Change()
  If isf <> 1 Then
    Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Text4.Text = HScroll3.Value
  End If
End Sub

Private Sub HScroll3_GotFocus()
  isf = 0
End Sub



Private Sub Picture1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  isf = 1
  GetCursorPos z    '得到当前点的坐标
  a = GetPixel(GetDC(0), z.X, z.Y)    '获得当前点的颜色
  B = a And &HFF     '分离出红色
  C = (a And 65280) \ 256    '分离出绿色
  d = (a And &HFF0000) \ 65536    '分离出蓝色
  If B < 16 Then    '分别判断各颜色的十进制值是否小于16，如果是就对各变量赋值0
  '目的是纠正可能是5个字符的情况，下同
     e = 0
     Else: e = ""
    End If
  If C < 16 Then
     F = 0
     Else: F = ""
    End If
  If d < 16 Then
     G = 0
     Else: G = ""
    End If
  
  HScroll1.Value = B
  HScroll2.Value = C
  HScroll3.Value = d
  Text2.Text = B
  Text3.Text = C
  Text4.Text = d
  Picture2.BackColor = RGB(B, C, d)
  'Label4.BackColor = RGB(B, C, d)
  Text1(0).Text = "(" & B & "," & C & "," & d & ")"
  Text1(1).Text = "#" & e & Hex(B) & F & Hex(C) & G & Hex(d)      '用VB的Hex函数将颜色值转换为16进制，并在Label2中显示
 End If
End Sub

Private Sub Text2_Change()
If isf <> 1 Then
  If Val(Text2.Text) <= 0 Then
    Text2.Text = 0
   ElseIf Val(Text2.Text) > 255 Then
    Text2.Text = 255
   ElseIf Val(Text2.Text) >= 0 And Val(Text2.Text) <= 255 Then
    Picture1.BackColor = RGB(Val(Text2.Text), Val(Text3.Text), Val(Text4.Text))
    HScroll1.Value = Val(Text2.Text)
  End If
  B = Val(Text2.Text)
  C = Val(Text3.Text)
  d = Val(Text4.Text)
  If B < 16 Then    '分别判断各颜色的十进制值是否小于16，如果是就对各变量赋值0
     e = 0
    Else: e = ""
    End If
  If C < 16 Then
     F = 0
    Else: F = ""
    End If
  If d < 16 Then
     G = 0
    Else: G = ""
    End If
  Picture1.BackColor = RGB(B, C, d)
  Me.BackColor = RGB(B, C, d)
  Text1(0).Text = "(" & B & "," & C & "," & d & ")"
  Text1(1).Text = "#" & e & Hex(B) & F & Hex(C) & G & Hex(d)
End If
End Sub

Private Sub Text2_GotFocus()
isf = 0
End Sub

Private Sub Text3_Change()
If isf <> 1 Then
  If Val(Text3.Text) <= 0 Then
    Text3.Text = 0
   ElseIf Val(Text3.Text) > 255 Then
    Text3.Text = 255
   ElseIf Val(Text3.Text) >= 0 And Val(Text3.Text) <= 255 Then
    Picture1.BackColor = RGB(Val(Text2.Text), Val(Text3.Text), Val(Text4.Text))
    HScroll2.Value = Val(Text3.Text)
  End If
  B = Val(Text2.Text)
  C = Val(Text3.Text)
  d = Val(Text4.Text)
  If B < 16 Then    '分别判断各颜色的十进制值是否小于16，如果是就对各变量赋值0
     e = 0
    Else: e = ""
    End If
  If C < 16 Then
     F = 0
    Else: F = ""
    End If
  If d < 16 Then
     G = 0
    Else: G = ""
    End If
  Picture1.BackColor = RGB(B, C, d)
  frmColor.BackColor = RGB(B, C, d)
  Text1(0).Text = "(" & B & "," & C & "," & d & ")"
  Text1(1).Text = "#" & e & Hex(B) & F & Hex(C) & G & Hex(d)
End If
End Sub

Private Sub Text3_GotFocus()
isf = 0
End Sub

Private Sub Text4_Change()
If isf <> 1 Then
  If Val(Text4.Text) <= 0 Then
    Text4.Text = 0
   ElseIf Val(Text4.Text) > 255 Then
    Text4.Text = 255
   ElseIf Val(Text4.Text) >= 0 And Val(Text4.Text) <= 255 Then
    Picture1.BackColor = RGB(Val(Text2.Text), Val(Text3.Text), Val(Text4.Text))
    HScroll3.Value = Val(Text4.Text)
  End If
  B = Val(Text2.Text)
  C = Val(Text3.Text)
  d = Val(Text4.Text)
  If B < 16 Then    '分别判断各颜色的十进制值是否小于16，如果是就对各变量赋值0
     e = 0
    Else: e = ""
    End If
  If C < 16 Then
     F = 0
    Else: F = ""
    End If
  If d < 16 Then
     G = 0
    Else: G = ""
    End If
  Picture1.BackColor = RGB(B, C, d)
  frmColor.BackColor = RGB(B, C, d)
  Text1(0).Text = "(" & B & "," & C & "," & d & ")"
  Text1(1).Text = "#" & e & Hex(B) & F & Hex(C) & G & Hex(d)
End If
End Sub

Private Sub Text4_GotFocus()
isf = 0
End Sub
