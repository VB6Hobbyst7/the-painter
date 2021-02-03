VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SProject 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Open"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin LP.Command Command3 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "SProject.frx":000C
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LP.Command Command1 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "SProject.frx":0028
      Enabled         =   0   'False
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LP.Command Command2 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1590
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "SProject.frx":0044
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1635
      Width           =   2415
   End
   Begin VB.CheckBox Start1 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   210
      TabIndex        =   0
      Top             =   3150
      Visible         =   0   'False
      Width           =   200
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LP.Command Command4 
      Height          =   360
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      IconAlign       =   1
      Icon            =   "SProject.frx":0060
      Caption         =   "  打开文件"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   3
      Begin LP.Command Command5 
         Height          =   315
         Left            =   4920
         TabIndex        =   13
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Icon            =   "SProject.frx":03B2
         Caption         =   "？"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLook      =   7
      End
      Begin VB.Image min 
         Height          =   315
         Left            =   5280
         Picture         =   "SProject.frx":03CE
         Top             =   30
         Width           =   315
      End
      Begin VB.Image quit 
         Height          =   315
         Left            =   5640
         Picture         =   "SProject.frx":0950
         Top             =   30
         Width           =   315
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000013&
      X1              =   399
      X2              =   399
      Y1              =   24
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000013&
      X1              =   400
      X2              =   0
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000013&
      X1              =   0
      X2              =   0
      Y1              =   24
      Y2              =   240
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3840
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "像素: 0000 X 0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2745
      TabIndex        =   6
      Top             =   2265
      Width           =   2190
   End
   Begin VB.Image Image3 
      Height          =   1515
      Left            =   435
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "请选择您需编辑的图片:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "下次启动程序时不启动本向导"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3150
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "下次启动程序时不启动本向导"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   495
      TabIndex        =   1
      Top             =   3165
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Image quit1 
      Height          =   315
      Left            =   3360
      Picture         =   "SProject.frx":0ED2
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image min0 
      Height          =   315
      Left            =   2280
      Picture         =   "SProject.frx":1454
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image quit0 
      Height          =   315
      Left            =   3000
      Picture         =   "SProject.frx":19D6
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image min1 
      Height          =   315
      Left            =   2640
      Picture         =   "SProject.frx":1F58
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "请选择您需编辑的图片:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1815
      TabIndex        =   5
      Top             =   975
      Width           =   2280
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "像素: 0000 X 0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "SProject.frx":24DA
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "SProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'扣色声明
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'完毕
'置顶声明
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'完毕
'在任务栏中显示无边框窗体的图标声明
''Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA " (ByVal hWnd As Long, ByVal nIndex As Long) As Long
''Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA " (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const GWL_STYLE = (-16)
'Private Const WS_SYSMENU = &H80000
'完毕




Sub Clear1()
min.Picture = min0.Picture
quit.Picture = quit0.Picture
End Sub

Private Sub Command1_Click()
Open "Temp.ini" For Output As #1
Write #1, Text1.Text
Close
FMain.Run1.Visible = True
FMain.Run2.Visible = False
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
cd1.Filter = lgT(3) & "|*.bmp;*.jpg;*.gif;*.wmf;*.ico"
cd1.flags = 2
cd1.ShowOpen
Text1.Text = cd1.FileName
If Text1.Text <> "" Then
Command1.Enabled = True
Label7.Visible = True
Label3.Visible = True
Image3.Visible = True
Label8.Visible = True
Image4.Picture = LoadPicture(Text1.Text)
Image3.Picture = LoadPicture(Text1.Text)
SetImage2 Image4.Width, Image4.Height

End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clear1
End Sub


Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\ap.htm", "", App.Path, 1

End Sub

Private Sub Form_Load()
'窗口
On Error Resume Next
'If Dir(App.Path & "\stepbg.skin") = "" Then
'MsgBox "Cannot find the Form Picture!", vbCritical, "Error"
'Unload Me
'End If

If LangA = "lgc" Then
Label6.Visible = False
Label4.Visible = False

End If


    rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)  '置顶
    
    '在任务栏中显示无边框窗体的图标
'Dim lStyle As Long

'lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_SYSMENU

'SetWindowLong hWnd, GWL_STYLE, lStyle ''
    '图标完毕
    
    
 '   Dim regWay As String
  '  regWay = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way")
  '  If regWay = "1" Then Start1.Value = 1
Label7.Visible = False
Label3.Visible = False
'Image3.Visible = False
'Label8.Visible = False


Label4.BackColor = RGB(250, 240, 204)
Label5(0).Caption = lgT(4)
Label5(1).Caption = Label5(0).Caption
Command1.Caption = "打开"
Command3.Caption = lgT(6)
 Label2 = lgT(7)
 Label1 = Label2
 Command2.Caption = lgT(375)
 FAbout.Hide
 Unload FAbout
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clear1
End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Label2_Click()
If Start1.Value = 0 Then
Start1.Value = 1
Else
Start1.Value = 0
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub min_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 1
min.Picture = min0.Picture
End Sub

Private Sub min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = min1.Picture
quit.Picture = quit0.Picture
End Sub

Private Sub quit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
quit.Picture = quit0.Picture
End Sub

Private Sub quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
quit.Picture = quit1.Picture
min.Picture = min0.Picture
End Sub


Private Sub Start1_Click()
 If Start1.Value = 0 Then
 Call DeleteValue(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way")
 ElseIf Start1.Value = 1 Then
 Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way", "1")
 End If
End Sub


Private Sub SetImage2(l%, B%)
Dim T As Single, newL%, newH%
Label7.Caption = lgT(2) & Format(l, "000") & " X " & Format(B, "000")
Label3.Caption = lgT(2) & Format(l, "000") & " X " & Format(B, "000")
newL = l: newH = B
T = 1
Do While newL > 100 Or newH > 100
newL = Int(l / T)
newH = Int(B / T)
T = T + 0.1
Loop
Image3.Width = newL
Image3.Height = newH
Image3.Move ((100 - newL) / 2) + 27, ((100 - newH) / 2) + 91

End Sub

