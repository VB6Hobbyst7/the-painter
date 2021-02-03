VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save As Options"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   Icon            =   "SaveAsOptions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2505
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   333
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   333
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2265
      Begin VB.OptionButton opt1Bit 
         Caption         =   "1 Bit - B/W"
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   350
         Width           =   2040
      End
      Begin VB.OptionButton opt4Bit 
         Caption         =   "4 Bit - 16 Colors"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   675
         Width           =   2025
      End
      Begin VB.OptionButton opt8Bit 
         Caption         =   "8 Bit - 256 Colors"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1035
         Width           =   2025
      End
      Begin VB.OptionButton opt24Bit 
         Caption         =   "24 Bit - True Color"
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   1395
         Value           =   -1  'True
         Width           =   2085
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'置顶声明
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'完毕


Private Sub cmdCancel_Click()
CancelIt = True
Unload Form2
End Sub

Private Sub cmdOK_Click()
Form2.Hide
End Sub

Private Sub Form_Load()
 rtn1 = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
If LangA = "lgc" Then
opt1Bit.Caption = "1 位 - 黑白"
opt4Bit.Caption = "4 位 - 16 色"
opt8Bit.Caption = "8 位 - 256 色"
opt24Bit.Caption = "24 位 - 真彩色"
cmdOK.Caption = "确定"
cmdCancel.Caption = "取消"
End If
End Sub
