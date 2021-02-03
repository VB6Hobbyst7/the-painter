VERSION 5.00
Begin VB.Form FLogo 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2040
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   5040
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   2865
      TabIndex        =   0
      Top             =   4785
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   6630
      Left            =   0
      Picture         =   "FIMLogo.frx":0000
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "FLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'置顶声明
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'完毕
'渐变声明
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'完毕
Dim alphaValue As Byte

Private Sub Form_Load()
 rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
Label2.Caption = "版权所有 (C) 2006 小画家工作室" & vbCrLf & vbCrLf & "版本: " & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "使用过程中需要帮助，请按 F1"

alphaValue = 250
End Sub



Sub log2()
FWel.Show
        If Command = "" Then
         ElseIf Command = "ap" Then
         Unload Me
           APrun
        ElseIf Command = "fp" Then
        Unload Me
             FPrun
        Else
        Unload Me
           ICONrun
        End If
End Sub


Private Sub Timer1_Timer()

On Error GoTo Skipnow
If alphaValue <= 0 Then

Timer3.Enabled = False
Unload Me

Else
  Dim rtn As Long
 rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
 rtn = rtn Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, rtn
  SetLayeredWindowAttributes hWnd, 0, alphaValue, LWA_ALPHA
  alphaValue = alphaValue - 16
End If
Exit Sub
Skipnow:

Unload Me

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer1.Enabled = True
log2
End Sub
