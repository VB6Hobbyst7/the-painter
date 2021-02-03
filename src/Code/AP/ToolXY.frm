VERSION 5.00
Begin VB.Form ToolXY 
   BorderStyle     =   0  'None
   Caption         =   "XY"
   ClientHeight    =   1125
   ClientLeft      =   75
   ClientTop       =   4110
   ClientWidth     =   4065
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   1125
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3480
   End
   Begin VB.Image Image21 
      Height          =   225
      Left            =   360
      Picture         =   "ToolXY.frx":0000
      Top             =   720
      Width           =   225
   End
   Begin VB.Image Image22 
      Height          =   225
      Left            =   1845
      Picture         =   "ToolXY.frx":0363
      Top             =   720
      Width           =   225
   End
   Begin VB.Image IList2 
      Height          =   210
      Left            =   2760
      Picture         =   "ToolXY.frx":06D3
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IList1 
      Height          =   210
      Left            =   2520
      Picture         =   "ToolXY.frx":09B5
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Ialpha2 
      Height          =   135
      Left            =   2040
      Picture         =   "ToolXY.frx":0C97
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Ialpha1 
      Height          =   135
      Left            =   1800
      Picture         =   "ToolXY.frx":0DD5
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   3720
      MouseIcon       =   "ToolXY.frx":0F13
      MousePointer    =   99  'Custom
      Picture         =   "ToolXY.frx":1065
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   120
      Picture         =   "ToolXY.frx":1347
      Top             =   100
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(295)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MouseIcon       =   "ToolXY.frx":1485
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Width           =   1440
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   855
      Left            =   4020
      Picture         =   "ToolXY.frx":15D7
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image11 
      Height          =   45
      Left            =   0
      Picture         =   "ToolXY.frx":1925
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4080
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   0
      MousePointer    =   15  'Size All
      Picture         =   "ToolXY.frx":22D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
   Begin VB.Image Image8 
      Enabled         =   0   'False
      Height          =   855
      Left            =   0
      Picture         =   "ToolXY.frx":6A3D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "ToolXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinPo As Boolean

Private Sub Form_Load()
 rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
 WinPo = True
 Me.Left = Screen.Width - Me.Width - 200
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Image4.Picture = IList1.Picture
End Sub

Private Sub Image2_DblClick()
Label3_Click
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Image4.Picture = IList1.Picture
End Sub

Private Sub Image4_Click()
ActiveTool = 1
PopupMenu FMain.mnuTool, , Image4.Left, Image4.Top + Image4.Height
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = IList2.Picture
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = IList1.Picture
End Sub




Private Sub Label3_Click()
If WinPo = True Then
Me.Height = 330
Image3.Picture = Ialpha2.Picture
WinPo = False
Else
Me.Height = 1125
Image3.Picture = Ialpha1.Picture
WinPo = True
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = True
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
End Sub


