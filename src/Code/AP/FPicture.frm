VERSION 5.00
Begin VB.Form FPicture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   ControlBox      =   0   'False
   Icon            =   "FPicture.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   135
      Width           =   1725
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   45
      TabIndex        =   6
      Top             =   480
      Width           =   1770
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3195
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      Top             =   3105
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   1845
      TabIndex        =   3
      Top             =   135
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "预览"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "撤消"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "平铺填充图片"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1365
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   3795
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   1
         Left            =   1395
         Max             =   10
         Min             =   1
         TabIndex        =   15
         Top             =   360
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "透明度"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   3105
         TabIndex        =   16
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "居中填充图片"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1365
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   3795
      Begin VB.OptionButton Option1 
         Caption         =   "按原大小居中"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   990
         Width           =   3330
      End
      Begin VB.OptionButton Option1 
         Caption         =   "拉伸至全屏"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   3330
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   0
         Left            =   1395
         Max             =   10
         Min             =   1
         TabIndex        =   9
         Top             =   360
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "透明度"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "大小: 000 X 000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   4770
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   1260
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   1500
   End
End
Attribute VB_Name = "FPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Sub Command1_Click() 'apply
FMain.Pic1.Picture = PicMem
SaveRedo
FMain.Pic1 = Im
FPicture.Hide
End Sub

Private Sub Command2_Click() 'cancel
FMain.Pic1.Picture = PicMem
FPicture.Hide
End Sub

Private Sub Command3_Click() 'show me
FMain.Pic1.Picture = PicMem
Screen.MousePointer = 11
Select Case Mix
Case 0 'mix picture
MixPic HScroll1(0).Value, Option1(0).Value
Case 1 'mix pattern
MixPattern HScroll1(1).Value
End Select
Command1.Enabled = True
Set Im = FMain.Pic1.Image
Screen.MousePointer = 1
End Sub

Private Sub Dir1_Change()
Command3.Enabled = False
On Error Resume Next
File1.Path = Dir1.Path
File1.Selected(0) = True
For Xx = 0 To File1.ListCount - 1
If File1.Selected(Xx) = True Then Command3.Enabled = True
Next Xx
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
    Pic2.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
Image1.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
DimensionImage Pic2.Width, Pic2.Height
End Sub

Private Sub Form_Activate()
'If LangA = "lge" Then
'FPictureen.Caption = Me.Caption
'Me.Hide
'FPictureen.Show 1

Label1(0).Caption = lgT(2) & " 000 X 000"
Frame1(1).Caption = lgT(229)
Frame1(0).Caption = lgT(359)
Label1(2).Caption = lgT(356)
Label1(1).Caption = lgT(356)
Option1(0).Caption = lgT(357)
Option1(1).Caption = lgT(358)
Command3.Caption = lgT(373)
Command1.Caption = lgT(360)
Command2.Caption = lgT(361)
'End If

On Error Resume Next
Command3.Enabled = False
Command1.Enabled = False
For Xx = 0 To 9
Frame1(Xx).Visible = False
Next Xx
Drive1.Enabled = False
Dir1.Enabled = False
'If Mix = 0 Then
Drive1.Enabled = True
Dir1.Enabled = True
'End If
If Mix = 1 Then
Dir1.Path = App.Path & "\Patterns"
End If
FPicture.Move FMain.Left + 200, FMain.Top + 3000, 4140, 7365
File1.Pattern = "*.bmp;*.gif;*.jpg;*.wmf"
File1.Selected(0) = True
For Xx = 0 To File1.ListCount - 1
If File1.Selected(Xx) = True Then Command3.Enabled = True
Next Xx
HScroll1(0).Value = 5
HScroll1(1).Value = 5
Set PicMem = FMain.Pic1.Image
Frame1(Mix).Visible = True
End Sub

Private Sub DimensionImage(l%, B%)
Dim T As Single, newL%, newH%

Label1(0).Caption = "大小: " & Format(l, "000") & " X " & Format(B, "000")
If LangA = "lge" Then Label1(0).Caption = lgT(2) & Format(l, "000") & " X " & Format(B, "000")
newL = l: newH = B
T = 1
Do While newL > 100 Or newH > 100
newL = Int(l / T)
newH = Int(B / T)
T = T + 0.1
Loop
Image1.Width = newL
Image1.Height = newH
Image1.Move ((100 - newL) / 2) + 88, ((100 - newH) / 2) + 210
End Sub

Private Sub Form_Load()


Image1.Move 88, 210
T3D FPicture, Image1, 5, T3dRaiseInset
Drive1.Drive = "C:\"
Dir1.Path = App.Path
File1.Path = App.Path
Option1(0).Value = True

End Sub

Private Sub HScroll1_Change(Index As Integer)
Label2(Index).Caption = Format(HScroll1(Index).Value / 10, "0.0")
End Sub

