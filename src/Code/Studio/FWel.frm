VERSION 5.00
Begin VB.Form FWel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "欢迎使用《小画家》"
   ClientHeight    =   5550
   ClientLeft      =   4590
   ClientTop       =   2700
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FWel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FWel.frx":0CCA
   ScaleHeight     =   5550
   ScaleWidth      =   7005
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   8760
      Width           =   180
   End
   Begin VB.CommandButton FCS 
      Caption         =   "Command4"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   25520
      Width           =   975
   End
   Begin LP.Command Command2 
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   2160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      Icon            =   "FWel.frx":3DC18
      Caption         =   "图片编辑"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   4
   End
   Begin LP.Command Command2 
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      Icon            =   "FWel.frx":3DC34
      Caption         =   "图标作坊"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   4
   End
   Begin LP.Command Command2 
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   4680
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      Icon            =   "FWel.frx":3DC50
      Caption         =   "取色吸管"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   4
   End
   Begin LP.Command Command2 
      Height          =   495
      Index           =   3
      Left            =   3840
      TabIndex        =   2
      Top             =   3000
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      Icon            =   "FWel.frx":3DC6C
      Caption         =   "涂鸦画板"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   4
   End
   Begin LP.Command Command2 
      Height          =   495
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   3840
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   873
      Icon            =   "FWel.frx":3DC88
      Caption         =   "快速抓图"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   4
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "获取屏幕上任意一点的RGB和Hex色值"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   15
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "方便快捷地抓取屏幕上的区域"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   14
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "建立和编辑图片文件"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   13
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ICON 图标文件编辑器"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "对现有图片进行美化处理"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   请在右侧的编辑器列表中单击您需要的编辑器。如需帮助，请按 F1。"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎使用《小画家》"
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
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   3360
      Picture         =   "FWel.frx":3DCA4
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   3240
      Picture         =   "FWel.frx":3E96E
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3480
      Picture         =   "FWel.frx":3F5B0
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
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
      Left            =   840
      TabIndex        =   0
      Top             =   4750
      Width           =   975
   End
   Begin VB.Image Bover 
      Height          =   300
      Left            =   600
      Picture         =   "FWel.frx":401F2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Bmove 
      Height          =   300
      Left            =   600
      Picture         =   "FWel.frx":419A4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image B1 
      Height          =   300
      Left            =   600
      Picture         =   "FWel.frx":43156
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   3000
      Picture         =   "FWel.frx":43D66
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   3120
      Picture         =   "FWel.frx":449A8
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   120
      Picture         =   "FWel.frx":455EA
      Top             =   120
      Width           =   6750
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
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
      Left            =   720
      TabIndex        =   1
      Top             =   4270
      Width           =   1215
   End
   Begin VB.Image I2 
      Height          =   300
      Left            =   600
      Picture         =   "FWel.frx":5B804
      Top             =   4200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image I1 
      Height          =   300
      Left            =   600
      Picture         =   "FWel.frx":5CFB6
      Top             =   4200
      Width           =   1500
   End
End
Attribute VB_Name = "FWel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
APrun
Case 1
Unload Me
ICONrun
Case 2
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "请重新执行安装文件，找不到plugin.exe", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "cs", App.Path, 1
End If
Case 3
FPrun
Case 4
QSGRun
End Select
End Sub


Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub

Private Sub Command2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub

Private Sub Form_Activate()
Move FWhole.Width / 2 - Me.Width / 2, FWhole.Height / 2 - Me.Height / 2 - 300
FCS.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
If FWhole.mAP.Checked = False And FWhole.mIC.Checked = False Then
MyMenu = GetSystemMenu(Me.hWnd, 0)
RemoveMenu MyMenu, &HF060, MF_BYCOMMAND
End If

'按钮锁定分析
If FWhole.mAP.Checked = False Then
Command2(0).Enabled = True
Else
Command2(0).Enabled = False
End If

If FWhole.mIC.Checked = False Then
Command2(1).Enabled = True
Else
Command2(1).Enabled = False
End If
'结束分析


Move FWhole.Width / 2 - Me.Width / 2, FWhole.Height / 2 - Me.Height / 2 - 300

regName = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic\Reginfo", "Show_ShKname")



If Label4.Caption = "" Then
Label4.Visible = False
Label6.Visible = False
Command3.Visible = False
End If

Label3.Caption = "帮助"
Label5.Caption = "设置"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = False
Bover.Visible = False
I2.Visible = False
FCS.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

IconTray = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "TrayIcon")
If IconTray = "" Then
Load frmTray
End If
End Sub

Private Sub I2_Click()
FSetting.Show 1
End Sub

Private Sub Image1_DblClick()
Text1.SetFocus
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove 0, 0, 1, 1
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove_MouseDown 0, 0, 5, 5
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
B1_MouseMove 0, 0, 5, 5
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove_MouseUp 0, 0, 5, 5
End Sub

Private Sub B1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = True
End Sub

Private Sub Bmove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bover.Visible = True
End Sub

Private Sub Bmove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu FWhole.mAbout, , B1.Left + B1.Width, B1.Top
Bmove.Visible = False
Bover.Visible = False
End Sub

Private Sub Label5_Click()
I2_Click
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
I1_MouseMove 0, 0, 5, 5
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i2_MouseUp 0, 0, 5, 5
End Sub

Private Sub I1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
I2.Visible = True
End Sub

Private Sub i2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
I2.Visible = False
I1.Visible = True
End Sub
