VERSION 5.00
Begin VB.Form FTimerQs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer Setting"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   Icon            =   "FTimerQs.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3285
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Caption         =   "Including the top forms"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   340
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Text            =   "100"
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Start to grab after:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Timer:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ms"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FTimerQs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long


Private Sub Command1_Click()
On Error GoTo err1
If Text1.Text <= 10000 And Text1.Text >= 0 Then
Me.Hide
Else: GoTo err1
End If
Exit Sub
err1:
MsgBox "Please input a value between 0 to 10000", vbInformation
End Sub

Private Sub Form_Load()
MyMenu = GetSystemMenu(Me.hWnd, 0)
RemoveMenu MyMenu, &HF060, MF_BYCOMMAND
 If LangA = "lgc" Then
 Me.Caption = "定时器设置"
 Label3.Caption = "请问在几毫秒后开始抓图？"
 Label2.Caption = "时间："
 Label1.Caption = "毫秒(ms)"
 Check1.Caption = "包括最顶层的窗体"
  End If
End Sub
