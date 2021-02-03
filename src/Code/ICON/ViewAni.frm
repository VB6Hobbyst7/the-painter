VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ViewAni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View .ani Files"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "ViewAni.frx":0000
   LinkTopic       =   "ViewAni"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4890
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   1200
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   1290
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Command1"
      Height          =   1395
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
End
Attribute VB_Name = "ViewAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function LoadCursor Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function LoadCursorFromFile Lib "User32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
Dim SystemCursor As Long
Dim vbCursor As Long
Dim inView As Boolean

Private Sub cmdOpen_Click()
Dim cdDir, cdIndex, Pos
On Error GoTo ex
  
'==========
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdViewAniDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdViewAniIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
'===========
NoRegVal: ComDia.CancelError = True
ComDia.FileName = ""
ComDia.flags = cdlOFNFileMustExist
ComDia.Filter = "Animated Cursors (*.ani)|*.ani"
ComDia.ShowOpen
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdViewAniDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdViewAniIndexSetting", cdIndex
'============
inView = True
vbCursor = LoadCursorFromFile(ComDia.FileName)
SetCapture Me.hWnd
 SetCursor vbCursor
Exit Sub
ex: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub Form_Click()
If inView Then
 ReleaseCapture
 SetCursor SystemCursor
 SystemCursor = 0
 inView = False
End If
End Sub

Private Sub Form_Load()
If LangA = "lgc" Then
cmdOpen.Caption = "打开一个 .ani 文件"
Label1.Caption = "注意：在“打开”对话框中请点击“打开”按钮（不要双击文件）"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub
