VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FText 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�������"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4590
   ControlBox      =   0   'False
   Icon            =   "FText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin LP.Command Command4 
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   4400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Icon            =   "FText.frx":000C
      Caption         =   "����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   7
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2280
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":0028
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":0318
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":0920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":0C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":1210
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FText.frx":1800
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "λ��"
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
      Height          =   1545
      Left            =   120
      TabIndex        =   13
      Top             =   2745
      Width           =   4335
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   720
         Left            =   3285
         TabIndex        =   31
         Top             =   675
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ButtonWidth     =   450
         ButtonHeight    =   423
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   855
         TabIndex        =   19
         Top             =   1080
         Width           =   1860
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   855
         TabIndex        =   17
         Top             =   810
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   285
         Index           =   2
         Left            =   3060
         TabIndex        =   16
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   285
         Index           =   1
         Left            =   1575
         TabIndex        =   15
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Y����"
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
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   22
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "X����"
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
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label6 
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
         Left            =   2745
         TabIndex        =   20
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label6 
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
         Left            =   2745
         TabIndex        =   18
         Top             =   765
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   4335
      Begin VB.Frame Frame3 
         Caption         =   "��Ӱ����"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   960
         Left            =   1710
         TabIndex        =   24
         Top             =   1125
         Width           =   2535
         Begin VB.HScrollBar HScroll3 
            Height          =   195
            Index           =   1
            Left            =   1080
            Max             =   20
            Min             =   -20
            TabIndex        =   26
            Top             =   585
            Value           =   1
            Width           =   870
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   195
            Index           =   0
            Left            =   1080
            Max             =   20
            Min             =   -20
            TabIndex        =   25
            Top             =   270
            Value           =   1
            Width           =   870
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "����Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   30
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "����X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   29
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00"
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
            Left            =   1980
            TabIndex        =   28
            Top             =   540
            Width           =   465
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00"
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
            Left            =   1980
            TabIndex        =   27
            Top             =   225
            Width           =   465
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FText.frx":1B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FText.frx":1C90
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FText.frx":1DEC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   180
         TabIndex        =   23
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyBold"
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyItalic"
               ImageIndex      =   2
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyUnderline"
               ImageIndex      =   3
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "������Ӱ"
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2160
         TabIndex        =   8
         Top             =   810
         Width           =   1500
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         LargeChange     =   10
         Left            =   2745
         Max             =   250
         Min             =   8
         TabIndex        =   5
         Top             =   405
         Value           =   8
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   45
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   315
         Width           =   2130
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Top             =   1665
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   11
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Lab3 
         Caption         =   "��Ӱ��ɫ"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1665
         Width           =   1185
      End
      Begin VB.Label Lab3 
         Caption         =   "������ɫ"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "��С"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2205
         TabIndex        =   7
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
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
         Left            =   3825
         TabIndex        =   6
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   165
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   135
      Width           =   4245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȷ��"
      Height          =   330
      Left            =   2640
      TabIndex        =   1
      Top             =   4410
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   330
      Left            =   3585
      TabIndex        =   0
      Top             =   4410
      Width           =   825
   End
End
Attribute VB_Name = "FText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Sub Check1_Click()
If Check1 = 0 Then
Frame3.Visible = False
FMain.Text(1).Visible = False
Else
Frame3.Visible = True
FMain.Text(1).Visible = True
End If
Text1.SetFocus
End Sub

Private Sub Combo1_Click()
FMain.Text(0).Font = Combo1.List(Combo1.ListIndex)
FMain.Text(1).Font = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
FMain.Text(0).Left = 0
HScroll2(0).Value = 0
Case 1
FMain.Text(0).Left = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
HScroll2(0).Value = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
Case 2
FMain.Text(0).Left = FMain.Pic1.Width - FMain.Text(0).Width
HScroll2(0).Value = FMain.Pic1.Width - FMain.Text(0).Width
End Select
End Sub

Private Sub Command2_Click()
SaveRedo
FMain.Pic1.Font = FMain.Text(0).Font
FMain.Pic1.FontSize = FMain.Text(0).FontSize
FMain.Pic1.FontBold = FMain.Text(0).FontBold
FMain.Pic1.FontItalic = FMain.Text(0).FontItalic
FMain.Pic1.FontUnderline = FMain.Text(0).FontUnderline
'set shadow first
If Check1.Value = 1 Then
FMain.Pic1.ForeColor = FMain.Text(1).ForeColor
FMain.Pic1.CurrentX = FMain.Text(1).Left
FMain.Pic1.CurrentY = FMain.Text(1).Top
FMain.Pic1.Print FMain.Text(1).Caption
End If
'now set text
FMain.Pic1.ForeColor = FMain.Text(0).ForeColor
FMain.Pic1.CurrentX = FMain.Text(0).Left
FMain.Pic1.CurrentY = FMain.Text(0).Top
FMain.Pic1.Print FMain.Text(0).Caption
For Xx = 0 To 1
FMain.Text(Xx).Visible = False
Next Xx
FText.Hide
End Sub

Private Sub Command3_Click() 'cancel
For Xx = 0 To 1
FMain.Text(Xx).Visible = False
Next Xx
FText.Hide
End Sub

Private Sub Command4_Click()
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\ap.htm", "", App.Path, 1

End Sub

Private Sub Form_Activate()
On Error Resume Next
FText.Move 30, 30
For Xx = 0 To 1
FMain.Text(Xx).Caption = Text1.Text
Next Xx
FMain.Text(0).ForeColor = Label3(0).BackColor
FMain.Text(1).ForeColor = Label3(1).BackColor
FMain.Text(0).Visible = True
Frame3.Visible = False
If Check1.Value = 1 Then
FMain.Text(1).Visible = True
Frame3.Visible = True
End If
HScroll2(0).Value = 0
HScroll2(0).min = -FMain.Pic1.Width
HScroll2(0).Max = FMain.Pic1.Width
HScroll2(0).Value = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
HScroll2(1).Value = 0
HScroll2(1).min = -FMain.Pic1.Height
HScroll2(1).Max = FMain.Pic1.Height
HScroll2(1).Value = (FMain.Pic1.Height - FMain.Text(0).Height) / 2
HScroll3(0).Value = 2
HScroll3(1).Value = 2
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)


Text1.Text = lgT(342)
Me.Caption = lgT(343)
Frame1.Caption = lgT(344)
Label2.Caption = lgT(345)
Check1.Caption = lgT(346)
Lab3(0).Caption = lgT(347)
Lab3(1).Caption = lgT(348)
Frame3.Caption = lgT(349)
Label10(0).Caption = lgT(350)
Label10(1).Caption = lgT(351)
Frame2.Caption = lgT(352)
Command1(0).Caption = lgT(353)
Command1(1).Caption = lgT(354)
Command1(2).Caption = lgT(355)
Label7(0).Caption = lgT(350)
Label7(1).Caption = lgT(351)
Command2.Caption = lgT(339)
Command3.Caption = lgT(340)


End Sub

Private Sub Form_Load()
On Error Resume Next
'FText.Move 0, 330, 4455, 5160


Combo1.Text = FMain.Text(0).Font
HScroll1.Value = 14
For Xx = 0 To 2
FMain.Text(Xx).FontSize = HScroll1.Value
Next Xx

End Sub

Private Sub HScroll1_Change()
On Error Resume Next
Label1.Caption = Format(HScroll1.Value, "000")
For Xx = 0 To 2
FMain.Text(Xx).FontSize = HScroll1.Value
Next Xx
Text1.SetFocus
End Sub

Private Sub HScroll2_Change(Index As Integer)
Label6(Index).Caption = HScroll2(Index).Value
FMain.Text(0).Left = HScroll2(0).Value
FMain.Text(0).Top = HScroll2(1).Value
Call SetShadow(HScroll3(0), HScroll3(1))
Text1.SetFocus
End Sub

Private Sub HScroll3_Change(Index As Integer)
Label9(Index).Caption = Format(HScroll3(Index).Value, "00")
Call SetShadow(HScroll3(0), HScroll3(1))
Text1.SetFocus
End Sub

Private Sub Label3_Click(Index As Integer)
On Error GoTo Label3Exit
Text1.SetFocus
FMain.CD1.flags = 3
FMain.CD1.Color = Label3(Index).BackColor
FMain.CD1.ShowColor
Label3(Index).BackColor = FMain.CD1.Color
FMain.Text(0).ForeColor = Label3(0).BackColor
FMain.Text(1).ForeColor = Label3(1).BackColor
Label3Exit:
End Sub

Private Sub Text1_Change()
FMain.Text(0).Caption = Text1.Text
FMain.Text(1).Caption = Text1.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Toolbar1.Buttons(1).Value = tbrPressed Then
FMain.Text(0).FontBold = True
FMain.Text(1).FontBold = True
Else
FMain.Text(0).FontBold = False
FMain.Text(1).FontBold = False
End If
If Toolbar1.Buttons(2).Value = tbrPressed Then
FMain.Text(0).FontItalic = True
FMain.Text(1).FontItalic = True
Else
FMain.Text(0).FontItalic = False
FMain.Text(1).FontItalic = False
End If
If Toolbar1.Buttons(3).Value = tbrPressed Then
FMain.Text(0).FontUnderline = True
FMain.Text(1).FontUnderline = True
Else
FMain.Text(0).FontUnderline = False
FMain.Text(1).FontUnderline = False
End If
Text1.SetFocus
End Sub

Private Sub SetShadow(dx%, dy%)
    FMain.Text(1).Left = FMain.Text(0).Left + dx
    FMain.Text(1).Top = FMain.Text(0).Top + dy
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
Case 1 'LU
HScroll2(0).Value = 0
HScroll2(1).Value = 0
Case 2 'CU
HScroll2(0).Value = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
HScroll2(1).Value = 0
Case 3 'RU
HScroll2(0).Value = FMain.Pic1.Width - FMain.Text(0).Width
HScroll2(1).Value = 0
Case 4 'LC
HScroll2(0).Value = 0
HScroll2(1).Value = (FMain.Pic1.Height - FMain.Text(0).Height) / 2
Case 5 'CC
HScroll2(0).Value = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
HScroll2(1).Value = (FMain.Pic1.Height - FMain.Text(0).Height) / 2
Case 6 'RC
HScroll2(0).Value = FMain.Pic1.Width - FMain.Text(0).Width
HScroll2(1).Value = (FMain.Pic1.Height - FMain.Text(0).Height) / 2
Case 7 'LD
HScroll2(0).Value = 0
HScroll2(1).Value = FMain.Pic1.Height - FMain.Text(0).Height
Case 8 'CD
HScroll2(0).Value = (FMain.Pic1.Width - FMain.Text(0).Width) / 2
HScroll2(1).Value = FMain.Pic1.Height - FMain.Text(0).Height
Case 9 'LD
HScroll2(0).Value = FMain.Pic1.Width - FMain.Text(0).Width
HScroll2(1).Value = FMain.Pic1.Height - FMain.Text(0).Height
End Select
End Sub
