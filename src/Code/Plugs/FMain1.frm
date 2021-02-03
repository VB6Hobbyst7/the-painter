VERSION 5.00
Begin VB.Form FMain1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magic Image Studio Plug-in List"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "FMain1.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4695
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Quick Screen Grab"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   1290
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Color Straw"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   330
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grab any part of your screen quickly and easily"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Get a color's RGB and Hex from any point of your screen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "FMain1.frx":0CCA
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   2
         Left            =   360
         Picture         =   "FMain1.frx":1994
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Plug-in List"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Image Studio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   825
      Left            =   240
      Picture         =   "FMain1.frx":25D6
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "FMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Command1_Click(Index As Integer)
Unload Me
Select Case Index
Case 0
frmColor.Show
Case 1
MDIForm1.Show
End Select
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE _
Or SWP_NOSIZE

If LangA = "lgc" Then
Label1.Caption = "插件列表"
Label3(0).Caption = "从您屏幕的任何一点上取得颜色的RGB和16进制Hex值"
Label3(1).Caption = "方便快速地抓取您屏幕的任何一部分"
End If

End Sub

