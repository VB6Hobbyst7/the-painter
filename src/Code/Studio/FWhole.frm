VERSION 5.00
Begin VB.MDIForm FWhole 
   BackColor       =   &H00808080&
   Caption         =   "LP"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9750
   Icon            =   "FWhole.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FWhole.frx":0CCA
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Menu mFile 
      Caption         =   "&Studio"
      Begin VB.Menu ms1 
         Caption         =   "Start!"
         Begin VB.Menu mIC 
            Caption         =   "图标作坊"
         End
         Begin VB.Menu mAP 
            Caption         =   "图片编辑"
         End
         Begin VB.Menu mFP 
            Caption         =   "涂鸦画板"
         End
         Begin VB.Menu mSG 
            Caption         =   "快速抓图"
         End
         Begin VB.Menu mCS 
            Caption         =   "取色吸管"
         End
      End
      Begin VB.Menu mj2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "&Help"
      Begin VB.Menu mHelpD 
         Caption         =   "Help Doc"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mj3 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout2 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FWhole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub LangSel()

'Lang Selection
mAbout2.Caption = "关于..."
mFile.Caption = "画室(&S)"
mAbout.Caption = "帮助(&H)"
ms1.Caption = "启动"
mExit.Caption = "退出"
Me.mHelpD.Caption = lgT(286)
End Sub


Private Sub mAbout2_Click()
about1.Show
End Sub

Private Sub mAM_Click()
With FAddPT
.Caption = "Add Plugins"
.Frame1(0).Visible = True
.Show 1
End With
End Sub

Private Sub mAP_Click()
APrun
End Sub

Private Sub mCS_Click()
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "Cannot Find Plugins Application", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "cs", App.Path, 1
End If
End Sub

Private Sub MDIForm_Load()
Me.Caption = lgT(404)
Me.BackColor = RGB(205, 117, 84)
Me.WindowState = 2
LangSel
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
iDone = True
End Sub



Private Sub MDIForm_Resize()
If Me.mAP.Checked = True Then
ToolsMini
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Unload Me
Unload frmTray
Unload FMain
ToolsClose
End
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mFP_Click()
FPrun
End Sub

Private Sub mHelpD_Click()
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\logo.htm", "", App.Path, 1

End Sub

Private Sub mIC_Click()
ICONrun
End Sub

Private Sub mPL_Click()
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "Cannot Find Plugins Application", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "", App.Path, 1
End If
End Sub


Private Sub mSG_Click()
QSGRun
End Sub

Private Sub mWeb_Click()
ShellExecute Me.hWnd, "Open", lgT(8), "", App.Path, 1
End Sub
