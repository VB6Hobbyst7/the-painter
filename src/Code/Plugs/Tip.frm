VERSION 5.00
Begin VB.Form Tip1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   Icon            =   "Tip.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4170
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "I know!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "in any situation."
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Press"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "To Grab the Screen and show the Main Window."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "But it's runing all the same."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Although the Main Window is hidden."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Tip1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 If LangA = "lgc" Then
 Me.Caption = "提示"
 Label1.Caption = "主窗口虽然隐藏"
 Label2.Caption = "但是进程仍然在后台运行中"
 Label3.Caption = "要 抓图 和 显示主窗口"
 Label4.Caption = "按下"
 Label6.Caption = "(在任何情况下)"
 Command1.Caption = "知道了!"
 
 End If
End Sub
