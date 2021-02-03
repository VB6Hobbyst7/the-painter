VERSION 5.00
Begin VB.Form about1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   3450
   ClientLeft      =   150
   ClientTop       =   12405
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3795
   StartUpPosition =   2  '屏幕中心
   Begin LP.Command Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "about.frx":000C
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "about.frx":0028
      Top             =   1440
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   1125
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   120
      Picture         =   "about.frx":0CF2
      Top             =   120
      Width           =   6750
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "建议至少使用 1024 * 768 分辨率"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "版权所有 (C) 2006 小画家工作室"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "about1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984


'置顶声明
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'完毕

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)

Label5.Caption = lgT(9)
Image2.Width = 3495
End Sub

Private Sub Label4_Click()

End Sub

