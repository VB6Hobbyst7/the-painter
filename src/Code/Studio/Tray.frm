VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3030
   ClientLeft      =   1545
   ClientTop       =   2130
   ClientWidth     =   4290
   Icon            =   "Tray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   4290
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   1560
      Top             =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "SysTray Popup Menu"
      Begin VB.Menu mnuREL 
         Caption         =   "还原编辑器列表"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nfIconData As NOTIFYICONDATA

Private Sub Form_Load()
On Error Resume Next
Timer1.Enabled = True

With nfIconData
 .hWnd = Me.hWnd
 .uID = Me.Icon
   .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.handle
 .szTip = "Magic Image Studio" & vbCrLf & "Version: " & ver1 & Chr$(0)
 .cbSize = Len(nfIconData)
 
          .dwState = 0
        .dwStateMask = 0
        
        
 End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)






End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
Case 7680 'MouseMove
Case 7695 'LeftMouseDown
mnuREL_Click
Case 7710 'LeftMouseUp
Case 7725 'LeftDblClick
Case 7740 'RightMouseDown
Case 7755 'RightMouseUp
 PopupMenu mnuPopup
Case 7770 'RightDblClick
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
Unload Me
End Sub




Private Sub mnuREL_Click()
On Error Resume Next
 FWel.Show
 FWel.SetFocus
 Unload Me
End Sub

Private Sub Timer1_Timer()
With nfIconData
 .hWnd = Me.hWnd
 .uID = Me.Icon
   .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.handle
 .szTip = "Magic Image Studio" & vbCrLf & "Version: " & ver1 & Chr$(0)
 .cbSize = Len(nfIconData)
 
          .dwState = 0
        .dwStateMask = 0
        If LangA = "lgc" Then
        .szInfoTitle = "显示编辑器列表点此图标" & Chr(0)
        .szInfo = "编辑器列表已经被隐藏" & vbCrLf & "要还原编辑器列表，请点此图标" & Chr(0)
        Else
        .szInfoTitle = "Click this icon for Editor List" & Chr(0)
        .szInfo = "The Editor List has now been hidden" & vbCrLf & "To restore the Editor List, Click this icon" & Chr(0)
        End If
        .dwInfoFlags = NIIF_INFO
         
       '  .uTimeout = 3
        
        
 End With
Call Shell_NotifyIcon(NIM_MODIFY, nfIconData)
Timer1.Enabled = False

End Sub
