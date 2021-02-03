VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import"
   ClientHeight    =   1125
   ClientLeft      =   975
   ClientTop       =   5970
   ClientWidth     =   6855
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   Begin VB.CommandButton Command1 
      Caption         =   "Import"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   315
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check Image to Import"
      Height          =   780
      Left            =   2070
      TabIndex        =   1
      Top             =   45
      Width           =   3525
      Begin VB.PictureBox picFill 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2565
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   225
         Width           =   510
      End
      Begin VB.PictureBox picRatio 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   225
         Width           =   510
      End
      Begin VB.OptionButton OptFill 
         Caption         =   "Fill"
         Height          =   540
         Left            =   1800
         TabIndex        =   3
         Top             =   190
         Width           =   735
      End
      Begin VB.OptionButton OptRatio 
         Caption         =   "Aspect Ratio"
         Height          =   540
         Left            =   120
         TabIndex        =   2
         Top             =   190
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Color RGB(197,197,197) will be imported as transparent."
      Height          =   195
      Left            =   1875
      TabIndex        =   7
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Image is larger than 32X32 pixels. Click and Drag to Select area to import."
      Height          =   990
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1770
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim w, H, r, C, pw, ph, pc
If AreaSel = False Then
If LangA = "lge" Then
MsgBox "Select an Area to Import."
Else
MsgBox "请选择要导入的区域"
End If
Exit Sub
End If
picRatio.Picture = picRatio.Image
picFill.Picture = picFill.Image

Me.MousePointer = 11
If IsCB = True Then
    If OptRatio Then
        For r = 0 To 31
        If w > 0 Then Exit For
            For C = 0 To 31
                If picRatio.Point(r, C) <> RGB(197, 197, 197) Then
                w = w + 1
                
                End If
            Next C
        Next r
        For C = 0 To 31
        If H > 0 Then Exit For
            For r = 0 To 31
                If picRatio.Point(r, C) <> RGB(197, 197, 197) Then
                H = H + 1
                End If
            Next r
        Next C
         Form1.picCB.Picture = LoadPicture()
         Form1.picCB.Width = w
         Form1.picCB.Height = H
         Form1.picMove.Width = w * 10
         Form1.picMove.Height = H * 10
         '=========
        For r = 0 To 31
            For C = 0 To 31
                If picRatio.Point(r, C) <> RGB(197, 197, 197) Then
                pc = picRatio.Point(r, C)
                Form1.picCB.PSet (pw, ph), pc
                pw = pw + 1
                End If
            Next C
            ph = ph + 1
            pw = 0
        Next r
    Else
         Form1.picCB.Width = 32
         Form1.picCB.Height = 32
         Form1.picMove.Height = 32
         Form1.picMove.Width = 32
         Form1.picCB.PaintPicture picFill.Picture, 0, 0, 32, 32
    End If

Else
    If OptRatio Then
         Form1.picReal.PaintPicture picRatio.Picture, 0, 0, 32, 32
    Else
         Form1.picReal.PaintPicture picFill.Picture, 0, 0, 32, 32
    End If
End If
Me.MousePointer = 0
iDone = True
IsCB = False
frmSelect.Hide
frmScroll.Hide
'Unload MDIForm1
End Sub

Private Sub Form_Load()
Me.Move 0, frmScroll.Height
'Me.ScaleWidth = frmScroll.ScaleWidth
'Me.Width = frmScroll.Width
picRatio.BackColor = RGB(197, 197, 197)
picFill.BackColor = RGB(197, 197, 197)
If LangA = "lgc" Then
Label1.Caption = "图像大小大于 32*32 像素，请圈选要导入的区域"
Frame1.Caption = "选择要导入的图像"
OptRatio.Caption = "原比例"
OptFill.Caption = "伸拉"
Command1.Caption = "导入"
Label2.Caption = "颜色RGB(197,197,197)导入时将会被当作透明色"

End If
End Sub

