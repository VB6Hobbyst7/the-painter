VERSION 5.00
Begin VB.Form frmLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Associate"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLink.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4185
   StartUpPosition =   2  '屏幕中心
   Begin LP.Command Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmLink.frx":000C
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
   Begin LP.XpRadioButton Option1 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Option1"
   End
   Begin LP.XpRadioButton Option1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Option1"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Don't ask me any more"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want to associate program with ico files"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
Call DeleteValue(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "IcoLink")
Else
Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "IcoLink", "1")
End If


End Sub

Private Sub Command1_Click()
'***Reg the file type****
If Option1(0).Value = True Then
Call savestring(HKEY_CLASSES_ROOT, "icofile\DefaultIcon", "", "%1")
Dim OpenPathV1 As String
 OpenPathV1 = App.Path & "\Studio.exe" & " /open %1"
Call savestring(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", "", OpenPathV1)
Call savestring(HKEY_CLASSES_ROOT, "icofile\Shell", "", "Open")
Call DeleteValue(HKEY_CLASSES_ROOT, ".ico", "PerceivedType")
End If
Unload Me
End Sub

Private Sub Form_Load()
If LangA = "lgc" Then
Label1.Caption = "是否将 ICO 文件关联至本编辑器？"
Option1(0).Caption = "是，关联! (推荐)"
Option1(1).Caption = "否"
Check1.Caption = "下次不再提示"
End If

End Sub

Private Sub Option1_Click(Index As Integer)
Option1(Index).Value = True
End Sub
