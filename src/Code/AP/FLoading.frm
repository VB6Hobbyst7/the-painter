VERSION 5.00
Begin VB.Form FLoading 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   495
   ClientLeft      =   2085
   ClientTop       =   1710
   ClientWidth     =   3120
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FLoading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Please wait...."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'ÖÃ¶¥ÉùÃ÷
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Íê±Ï

Private Sub Form_Load()
Label9.Caption = lgT(327)

   rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)  'ÖÃ¶¥
End Sub

Private Sub Form_Resize()
With Me
.Height = 570
.Width = 3240
End With
End Sub
