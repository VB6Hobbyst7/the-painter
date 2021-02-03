VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5655
   StartUpPosition =   2  '屏幕中心
   Begin LP.Command Command3 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Icon            =   "Form5.frx":000C
      Caption         =   "帮助"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   7
   End
   Begin VB.Frame Frame1 
      Caption         =   "图标作坊"
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   5415
      Begin LP.XpRadioButton Option2 
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option1"
      End
      Begin LP.XpRadioButton Option2 
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   14
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Value           =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option1"
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Associate program with ico files"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
   End
   Begin LP.Command Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "Form5.frx":0028
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LP.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "Form5.frx":0044
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1080
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "全局"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox Check1 
         Caption         =   "Display the Sales Promotion at Studio startup"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enable the tray icon for restoreing the Editor List"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "图片编辑"
      Height          =   1935
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5415
      Begin VB.CheckBox Check4 
         Caption         =   "Enable the Start Guide"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   4335
         Begin VB.OptionButton Option1 
            Caption         =   "Default"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   3615
         End
         Begin LP.Command Command2 
            Height          =   255
            Left            =   3150
            TabIndex        =   11
            Top             =   620
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Icon            =   "Form5.frx":0060
            Caption         =   "XpButton"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Custom"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Enable the background (Possess Physical Memory)"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check3_Click()
If Check3.Value = 0 Then
Option1(0).Enabled = False
Option1(1).Enabled = False
Text1.Enabled = False
Command2.Enabled = False
Else
Option1(0).Enabled = True
Option1(1).Enabled = True
Option1(0).Value = True
End If
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
        'Save Promotion
        If Check1.Value = 1 Then
        Call DeleteValue(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "BuyForm")
        Else
        Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "BuyForm", "1")
        End If
        'Save Tray Icon
        If Check2.Value = 1 Then
        Call DeleteValue(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "TrayIcon")
        Else
        Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "TrayIcon", "1")
        End If
        'Save AP bg
                 If Check3.Value = 1 And Option1(1).Value = True Then
                      If Text1.Text = "" Then
                              If LangA = "lgc" Then
                              MsgBox "背景图片未选择", vbExclamation, "Error"
                              Else
                              MsgBox "Select the Background picture please.", vbExclamation, "Error"
                              End If
                      For i = 0 To 3
                      Frame1(i).Visible = False
                      Next
                      Frame1(2).Visible = True
                      Command2.SetFocus
                      Exit Sub
                      End If
                 End If
           If Check3.Value = 1 Then
                  If Option1(0).Value = True Then
                      Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Apbg", "")
                  ElseIf Option1(1).Value = True Then
                      Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Apbg", "1")
                      Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Apbgadd", Text1.Text)
                  End If
           ElseIf Check3.Value = 0 Then ' None
                      Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "APbg", "0")
           End If
        'Save Start Guide
        If Check4.Value = 0 Then
        Call DeleteValue(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way")
        Else
        Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way", "1")
        End If
        'ICON Type
        If Option2(0).Value = True Then
        Call savestring(HKEY_CLASSES_ROOT, "icofile\DefaultIcon", "", "%1")
        Dim OpenPathV1 As String
        OpenPathV1 = App.Path & "\Studio.exe" & " /open %1"
        Call savestring(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", "", OpenPathV1)
        Call savestring(HKEY_CLASSES_ROOT, "icofile\Shell", "", "Open")
        Call DeleteValue(HKEY_CLASSES_ROOT, ".ico", "PerceivedType")
        Else
        Call DeleteValue(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", "")

        End If



End If
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
CD1.Filter = lgT(3) & "|*.bmp;*.jpg;*.gif;*.wmf;*.ico"
CD1.flags = 2
CD1.ShowOpen
If CD1.FileName <> "" Then Text1.Text = CD1.FileName

End Sub




Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\main.htm", "", App.Path, 1

End Sub

Private Sub Form_Load()
On Error Resume Next

Command1(0).SetFocus
' Dialog Language
If LangA = "lgc" Then
Check2.Caption = "使用通知栏小图标，用于唤出编辑器列表"
Check3.Caption = "使用背景图片 (会占用内存)"
Option1(0).Caption = "默认背景"
Option1(1).Caption = "自定义"
Command2.Caption = "浏览..."
Command1(0).Caption = "确定(&O)"
Command1(1).Caption = "取消(&C)"
'TabStrip1.Tabs(1).Caption = "画室全局"
Me.Caption = "参数设置"
Label5.Caption = "是否将 ICO 文件关联至《小画家》?"
Option2(0).Caption = "是 (推荐)"
Option2(1).Caption = "否"
Check4.Caption = "启动时直接切换至效果预览式编辑"
End If

' Sales Promotion
BuyFormShow = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "BuyForm")
If BuyFormShow = "" Then Check1.Value = 1
' Icon Tray
IconTray = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "TrayIcon")
If IconTray = "" Then Check2.Value = 1



'Amazing Picture

' Bg
apbg = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "APbg")
apbgadd = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "APbgadd")
If apbg = "" Then ' Default
Check3.Value = 1
Option1(0).Value = True
Text1.Enabled = False
Command2.Enabled = False
ElseIf apbg = "0" Then ' None
Check3.Value = 0
Option1(0).Enabled = False
Option1(1).Enabled = False
Text1.Enabled = False
Command2.Enabled = False
ElseIf apbg = "1" Then ' Custom
Check3.Value = 1
Option1(1).Value = True
Text1.Text = apbgadd
End If
'Start Guide
regWay = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way")
If regWay = "1" Then
Check4.Value = 1
End If
'ICON Type
Dim GetV1, GetV1set, OpenPathV1 As String
 OpenPathV1 = App.Path & "\Studio.exe" & " /open %1"
GetV1 = getstring(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", "")
If GetV1 = OpenPathV1 Then
Option2(0).Value = True
Option2(1).Value = False
End If
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 1 Then
Text1.Enabled = True
Command2.Enabled = True
Else
Text1.Enabled = False
Command2.Enabled = False
End If
End Sub
