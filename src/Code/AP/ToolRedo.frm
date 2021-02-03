VERSION 5.00
Begin VB.Form ToolRedo 
   BorderStyle     =   0  'None
   Caption         =   "Redo"
   ClientHeight    =   1245
   ClientLeft      =   75
   ClientTop       =   1170
   ClientWidth     =   4065
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.Image IList2 
      Height          =   210
      Left            =   2760
      Picture         =   "ToolRedo.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IList1 
      Height          =   210
      Left            =   2520
      Picture         =   "ToolRedo.frx":02E2
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Ialpha2 
      Height          =   135
      Left            =   2040
      Picture         =   "ToolRedo.frx":05C4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Ialpha1 
      Height          =   135
      Left            =   1800
      Picture         =   "ToolRedo.frx":0702
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   3720
      MouseIcon       =   "ToolRedo.frx":0840
      MousePointer    =   99  'Custom
      Picture         =   "ToolRedo.frx":0992
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   120
      Picture         =   "ToolRedo.frx":0C74
      Top             =   100
      Width           =   135
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   60
      Stretch         =   -1  'True
      Top             =   375
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   855
      Stretch         =   -1  'True
      Top             =   375
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   1650
      Stretch         =   -1  'True
      Top             =   375
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   2445
      Stretch         =   -1  'True
      Top             =   375
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   4
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   375
      Width           =   735
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
      MouseIcon       =   "ToolRedo.frx":0DB2
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Width           =   1440
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   975
      Left            =   4025
      Picture         =   "ToolRedo.frx":0F04
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image11 
      Height          =   45
      Left            =   0
      Picture         =   "ToolRedo.frx":1252
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4080
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   0
      MousePointer    =   15  'Size All
      Picture         =   "ToolRedo.frx":1C00
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
   Begin VB.Image Image8 
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      Picture         =   "ToolRedo.frx":636A
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "ToolRedo"
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

Private Sub Image1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
Redo
ElseIf Index = 1 Then
Redo
Redo
ElseIf Index = 2 Then
Redo
Redo
Redo
ElseIf Index = 3 Then
Redo
Redo
Redo
Redo
ElseIf Index = 4 Then
Redo
Redo
Redo
Redo
Redo
End If
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
ActiveTool = 0
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
Me.Height = 1245
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
