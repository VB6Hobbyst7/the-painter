VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Grab"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   945
   ClientWidth     =   9120
   Icon            =   "ScreenGrabMDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "ScreenGrabMDIForm1.frx":0CCA
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   2400
      Top             =   3840
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   3  'Not Merge Pen
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "ScreenGrabMDIForm1.frx":AC90C
      ScaleHeight     =   375
      ScaleWidth      =   9120
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      Begin VB.CommandButton Command1 
         Caption         =   "Hide Window"
         Height          =   320
         Left            =   6360
         TabIndex        =   1
         Top             =   30
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1080
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Picture Properties"
      End
      Begin VB.Menu j1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseChild 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   " -= &Capture =-"
      Begin VB.Menu grabnow 
         Caption         =   "&Grab Screen Now!"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGrabScreen 
         Caption         =   "Grab &Screen after..."
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu hp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu hp2 
         Caption         =   "帮助..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETHOTKEY = &H32
Private Const HOTKEYF_SHIFT = &H1
Private Const HOTKEYF_CONTROL = &H2
Private Const HOTKEYF_ALT = &H4
Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer


'检测到定义的快捷键被按下
Private Function MyHotKey(vKeyCode As Integer) As Boolean
  MyHotKey = (GetAsyncKeyState(vKeyCode) < 0)
End Function



Function GetScreenSnapshot(Optional ByVal hWnd As Long) As IPictureDisp

    Dim targetDC As Long
    Dim hdc As Long
    Dim tempPict As Long
    Dim oldPict As Long
    Dim wndWidth As Long
    Dim wndHeight As Long
    Dim Pic As PICTDESC
    Dim rcWindow As RECT
    Dim guid(3) As Long

    ' provide the right handle for the desktop window

    If hWnd = 0 Then hWnd = GetDesktopWindow

    ' get window's size
    GetWindowRect hWnd, rcWindow
    wndWidth = rcWindow.Right - rcWindow.Left
    wndHeight = rcWindow.Bottom - rcWindow.Top
    ' get window's device context
    targetDC = GetWindowDC(hWnd)

    ' create a compatible DC
    hdc = CreateCompatibleDC(targetDC)

    ' create a memory bitmap in the DC just created
    ' the has the size of the window we're capturing
    tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
    oldPict = SelectObject(hdc, tempPict)

    ' copy the screen image into the DC
    BitBlt hdc, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy

    ' set the old DC image and release the DC
    tempPict = SelectObject(hdc, oldPict)
    DeleteDC hdc
    ReleaseDC GetDesktopWindow, targetDC

    ' fill the ScreenPic structure

    With Pic

        .cbSize = Len(Pic)
        .pictType = 1           ' means picture
        .hIcon = tempPict
        .hPal = 0           ' (you can omit this of course)

    End With

    ' convert the image to a IpictureDisp object
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect Pic, guid(0), True, GetScreenSnapshot

End Function

Private Sub Command1_Click()
Tip1.Show 1
Me.Hide
End Sub

Private Sub grabnow_Click()
GrabNow1
End Sub

Private Sub hp2_Click()
ShellExecute Me.hWnd, "Open", App.Path + "\Help\grab.htm", "", App.Path, 1
End Sub

Private Sub MDIForm_Load()
   Dim l As Long
   Dim wHotkey As Long
   wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * (2 ^ 8) + 65
   l = SendMessage(FTimerQs.hWnd, WM_SETHOTKEY, wHotkey, 0)
Me.Caption = "快速抓图 - 小画家"
Me.BackColor = RGB(255, 126, 176)
Picture1.BackColor = RGB(166, 207, 251)
Picture1.PSet (600, 100)
Picture1.Print "任何情况下按下 [ F4 ] 直接抓图."
mnuFile.Caption = "文件(&F)"
mnuSave.Caption = "保存"
Command1.Caption = "隐藏本窗"
mnuProperties.Caption = "图片属性"
mnuCloseChild.Caption = "关闭"
mnuExit.Caption = "退出"
mnuCapture.Caption = " -= 抓图 =- (F4)"
grabnow.Caption = "立即抓图"
mnuGrabScreen.Caption = "定时抓图..."

End Sub

Private Sub MDIForm_Resize()
Command1.Move Me.Width - 1500

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If debugme = True Then MsgBox "unload event"

    If debugme = True Then MsgBox MDIForm1.ActiveForm Is Nothing
    
    Unload Form1
    Unload Me
End
End Sub

Private Sub mnuAbout_Click()
about1.Show 1
End Sub

Private Sub mnuCloseChild_Click()

    If Not (ActiveForm Is Nothing) Then Unload ActiveForm

End Sub

Private Sub mnuExit_Click()
    'child forms take care of themselves
    Unload Form1
    Unload Me
End Sub

Private Sub mnuGrabScreen_Click()

    '
    FTimerQs.Show 1
    MDIForm1.Visible = False
    Form1.Visible = False
    'start main grab sequence
    Timer1.Interval = FTimerQs.Text1.Text
    Timer1.Enabled = True

End Sub

'Private Sub mnuNewChildForm_Click()
''    Dim frmChild As New frmChild
''    frmChild.Show
''    MsgBox frmChild.Width & ":" & frmChild.Height
''    MsgBox frmChild.ScaleWidth & ":" & frmChild.ScaleHeight
'End Sub

Private Sub mnuProperties_Click()

    If MDIForm1.ActiveForm Is Nothing Then Exit Sub

    With MDIForm1.ActiveForm.Picture1
        'MsgBox "Picture Width= " & MDIForm1.ActiveForm.Picture1.Picture.Width & ": Height=" & MDIForm1.ActiveForm.Picture1.Picture.Height
         If LangA = "lge" Then
         MsgBox "Picture width= " & CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels)) _
           & "  :  Picture height= " & CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels)), vbInformation
         Else
        MsgBox "图像宽度= " & CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels)) _
           & "  :  图像高度= " & CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels)), vbInformation
           End If
    End With

End Sub

Private Sub mnuSave_Click()
    '
    'MsgBox "MDIForm1.ActiveForm.IsDirty=" & MDIForm1.ActiveForm.IsDirty
    If MDIForm1.ActiveForm Is Nothing Then Exit Sub
    'no need to save
    If MDIForm1.ActiveForm.IsDirty = False Then Exit Sub

    If savepictureRoutine = True Then
        'reset IsDirty flag
        MDIForm1.ActiveForm.IsDirty = False
        'update menu
        MDIForm1.mnuSave.Enabled = False
    Else
        'no change of isdirty flag/property
    End If

End Sub

Private Sub Timer1_Timer()
On Error GoTo err1
    Static count
    ''    MsgBox count
    'wait a bit then get screen;
    'count 0-4 ;timer on 500 interval
    If count > 3 Then

        Form1.Picture = GetScreenSnapshot(0)
        count = 0
        Timer1.Enabled = False
        Form1.Visible = True

    End If

    count = count + 1
    Exit Sub
err1:
MsgBox "Error", vbCritical, "Error"
End Sub

Private Sub Timer2_Timer()
 If MyHotKey(vbKeyF4) Then
GrabNow1
End If
End Sub


Sub GrabNow1()
    MDIForm1.Visible = False
    Form1.Visible = False
    'start main grab sequence
    Timer1.Interval = 20
    Timer1.Enabled = True
End Sub
