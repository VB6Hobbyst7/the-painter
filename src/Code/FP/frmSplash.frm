VERSION 5.00
Begin VB.Form FrmSplash 
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "正在启动，请稍候..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2715
   End
End
Attribute VB_Name = "frmSPlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
x = SetWindowPos(frmSPlash.hWnd, -1, 0, 0, 0, 0, 3)
alphaValue = 250
End Sub

