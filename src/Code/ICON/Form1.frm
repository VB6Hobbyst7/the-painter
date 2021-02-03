VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H009E4D3F&
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9195
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "Form1.frx":0CCE
   Picture         =   "Form1.frx":1110
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   Begin VB.CommandButton cmdMore 
      Caption         =   "More color"
      Height          =   615
      Left            =   7320
      TabIndex        =   52
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton FCS 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1200
      TabIndex        =   51
      Top             =   1.71720e5
      Width           =   495
   End
   Begin LP.Command cmdPicker 
      Height          =   375
      Left            =   6480
      TabIndex        =   50
      Top             =   6000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Icon            =   "Form1.frx":1E68
      Caption         =   "XpButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLook      =   15
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyNew"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyOpen"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyOpen1"
                  Object.Tag             =   "keyOpen1"
                  Text            =   "Op"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyOpen2"
                  Object.Tag             =   "keyOpen2"
                  Text            =   "&Extract Icon from a file"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySave"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyUndo"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyRedo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyCut"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyCopy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyPaste"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyDel"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyCS"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyPlugin"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyCS2"
                  Text            =   "取色吸管"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCBsprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6105
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   44
      Top             =   5505
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picCBmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6105
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   43
      Top             =   4770
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picCB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6135
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   41
      Top             =   4185
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   7875
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7335
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "Redo"
      Height          =   390
      Left            =   210
      TabIndex        =   39
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   390
      Left            =   210
      TabIndex        =   38
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6240
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   5715
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      DragIcon        =   "Form1.frx":1E84
      Height          =   480
      Left            =   5145
      LinkTimeout     =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picReal16 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      Top             =   960
      Width           =   240
   End
   Begin VB.PictureBox picHand 
      Height          =   420
      Left            =   9240
      Picture         =   "Form1.frx":1FD6
      ScaleHeight     =   360
      ScaleWidth      =   405
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton cmdRegion 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   570
      Picture         =   "Form1.frx":2128
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Select Region & Clipboard Functions"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   30
      Top             =   5685
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picA 
      Height          =   375
      Left            =   9240
      Picture         =   "Form1.frx":2442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Text"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picFlood 
      Height          =   375
      Left            =   9600
      Picture         =   "Form1.frx":2594
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdCircleDraw 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   570
      Picture         =   "Form1.frx":26E6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Ellipse"
      Top             =   2310
      Width           =   330
   End
   Begin VB.CommandButton cmdFillCircleDraw 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   165
      Picture         =   "Form1.frx":29CC
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Filled Ellipse"
      Top             =   2310
      Width           =   330
   End
   Begin VB.CommandButton cmdFillBox 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   165
      Picture         =   "Form1.frx":2CB2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Filled Rectangle"
      Top             =   2715
      Width           =   330
   End
   Begin VB.CommandButton cmdRect 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   570
      Picture         =   "Form1.frx":2FCC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Rectangle"
      Top             =   2715
      Width           =   330
   End
   Begin VB.CommandButton cmdLine 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   165
      Picture         =   "Form1.frx":32B2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Line"
      Top             =   1905
      Width           =   330
   End
   Begin VB.PictureBox picBasic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   16
      Height          =   1710
      Left            =   6420
      MousePointer    =   99  'Custom
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   17
      Top             =   4005
      Width           =   2250
   End
   Begin VB.PictureBox picColorpicker 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":3674
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   6270
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picEraser 
      Height          =   240
      Left            =   135
      Picture         =   "Form1.frx":37C6
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   6435
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picPencil 
      Height          =   330
      Left            =   150
      Picture         =   "Form1.frx":3918
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   6075
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTransparent 
      Height          =   285
      Left            =   6720
      ScaleHeight     =   225
      ScaleWidth      =   270
      TabIndex        =   12
      Top             =   825
      Width           =   330
   End
   Begin VB.PictureBox pic16Color 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   16
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   645
      Left            =   6435
      MousePointer    =   99  'Custom
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   11
      Top             =   3000
      Width           =   2250
   End
   Begin VB.CommandButton cmdErase 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   570
      Picture         =   "Form1.frx":3A6A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eraser"
      Top             =   1500
      Width           =   330
   End
   Begin VB.CommandButton cmdFlood 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   570
      Picture         =   "Form1.frx":3D4C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Flood"
      Top             =   1905
      Width           =   330
   End
   Begin VB.CommandButton cmdPaint 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   165
      Picture         =   "Form1.frx":3DE9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pixel Paint"
      Top             =   1500
      Width           =   330
   End
   Begin VB.PictureBox PicMseColor 
      BackColor       =   &H000000FF&
      Height          =   525
      Left            =   8040
      MouseIcon       =   "Form1.frx":4007
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   540
   End
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   480
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picReal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      ToolTipText     =   "Clik for 16x16 view"
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox PicMseColorR 
      BackColor       =   &H0000FFFF&
      Height          =   540
      Left            =   8400
      MouseIcon       =   "Form1.frx":4159
      MousePointer    =   99  'Custom
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   1680
      Width           =   540
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1155
      MouseIcon       =   "Form1.frx":42AB
      MousePointer    =   99  'Custom
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      Top             =   1470
      Width           =   4815
      Begin VB.PictureBox picMove 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2220
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   42
         Top             =   1290
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picTest 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   765
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   690
      End
      Begin PicClip.PictureClip PicClipMove 
         Left            =   450
         Top             =   1380
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   393216
      End
      Begin PicClip.PictureClip picClip 
         Left            =   0
         Top             =   0
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   393216
      End
      Begin VB.Shape shRect 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   420
         Left            =   765
         Top             =   945
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   10
         FillColor       =   &H00C0E0FF&
         Height          =   420
         Left            =   90
         Shape           =   2  'Oval
         Top             =   1755
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   123
         X2              =   183
         Y1              =   12
         Y2              =   12
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":43FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":450F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4621
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4733
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4845
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4957
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A69
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4C8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4EFF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6300
      TabIndex        =   53
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   49
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   48
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   1200
      MouseIcon       =   "Form1.frx":5011
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   975
      Width           =   2055
   End
   Begin VB.Image imgSaveDown 
      Height          =   360
      Left            =   960
      Picture         =   "Form1.frx":531B
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSaveOver 
      Height          =   360
      Left            =   1200
      Picture         =   "Form1.frx":5A1D
      ToolTipText     =   "Save As"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpenDown 
      Height          =   360
      Left            =   1440
      Picture         =   "Form1.frx":611F
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpenOver 
      Height          =   360
      Left            =   960
      Picture         =   "Form1.frx":6821
      ToolTipText     =   "Open"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNewDown 
      Height          =   360
      Left            =   240
      Picture         =   "Form1.frx":6F23
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNewOver 
      Height          =   360
      Left            =   600
      Picture         =   "Form1.frx":75A5
      ToolTipText     =   "New"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblMsePos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   2190
   End
   Begin VB.Shape shSel 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      Height          =   405
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1455
      Width           =   405
   End
   Begin VB.Image imgSave 
      Height          =   360
      Left            =   1560
      Picture         =   "Form1.frx":7C27
      ToolTipText     =   "Save As"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpen 
      Height          =   360
      Left            =   600
      Picture         =   "Form1.frx":7DC9
      ToolTipText     =   "Open"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNew 
      Height          =   345
      Left            =   210
      Picture         =   "Form1.frx":7F6B
      ToolTipText     =   "New"
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblRGB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Left            =   6000
      TabIndex        =   20
      Top             =   6480
      Width           =   2940
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   6285
      TabIndex        =   19
      Top             =   2310
      Width           =   2715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   1
      Left            =   6495
      TabIndex        =   18
      Top             =   3795
      Width           =   2040
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   630
      Left            =   5880
      TabIndex        =   13
      Top             =   870
      Width           =   3090
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "The 16 Named Colors."
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Index           =   0
      Left            =   6465
      TabIndex        =   10
      Top             =   2760
      Width           =   2400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   6480
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Left            =   840
      TabIndex        =   8
      Top             =   720
      Width           =   1770
   End
   Begin VB.Line Line1 
      X1              =   16
      X2              =   616
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5985
      Picture         =   "Form1.frx":85D5
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lable1"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   6285
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   7
      Height          =   4935
      Left            =   1110
      Top             =   1410
      Width           =   4920
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   840
      TabIndex        =   45
      Top             =   6720
      Width           =   7575
   End
   Begin VB.Image Bover 
      Height          =   300
      Left            =   3480
      Picture         =   "Form1.frx":9217
      Top             =   840
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Bmove 
      Height          =   300
      Left            =   3480
      Picture         =   "Form1.frx":A9C9
      Top             =   840
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image B1 
      Height          =   300
      Left            =   3480
      Picture         =   "Form1.frx":C17B
      Top             =   840
      Width           =   1500
   End
   Begin VB.Image Bover2 
      Height          =   300
      Left            =   30
      Picture         =   "Form1.frx":CD8B
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Bmove2 
      Height          =   300
      Left            =   30
      Picture         =   "Form1.frx":E53D
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image b2 
      Height          =   300
      Left            =   30
      Picture         =   "Form1.frx":FCEF
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&F"
      Begin VB.Menu mnuNew 
         Caption         =   "&N"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&O"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&S"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&E"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&E"
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&U"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&R"
         Index           =   1
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "Cu&t"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&C"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&P"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&D"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "C&le"
         Index           =   8
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "C&an"
         Index           =   9
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Options"
      Begin VB.Menu mnuExtract 
         Caption         =   "&Ex"
      End
      Begin VB.Menu mnuAni 
         Caption         =   "Cre"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu nmuAni 
         Caption         =   "&Vi"
      End
      Begin VB.Menu mnuChgPix 
         Caption         =   "&Ch"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAssoc 
         Caption         =   "A&sso"
      End
   End
   Begin VB.Menu nmuEffects 
      Caption         =   "E&ff"
      Begin VB.Menu mnuHorz 
         Caption         =   "Flip &H"
      End
      Begin VB.Menu mnuVert 
         Caption         =   "Flip &V"
      End
      Begin VB.Menu mnuRotateRight 
         Caption         =   "90 Light"
      End
      Begin VB.Menu mnuRotateLeft 
         Caption         =   "90 &Left"
      End
   End
   Begin VB.Menu mnuAbout2 
      Caption         =   "&H"
      Begin VB.Menu mnuHelp 
         Caption         =   "帮助..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&C"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "&P"
         Visible         =   0   'False
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&D"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "C&an"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim okToMove As Boolean, prevX%, prevY%
Dim xSav, ySav, xStart, yStart
Dim pX%, pY%, pXOff%, pYOff%
'==========
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'==========
Private Declare Function DestroyIcon Lib "User32" (ByVal hIcon As Long) As Long
Private CurrentFile$
Private CurrentName$
Dim cvtY, cvtX
Dim filePath
Dim pixelDraw As Boolean, canDraw As Boolean ', Dirty As Boolean
Dim ColorChg, bkClr, clr As Long, r As Integer, G As Integer, B As Integer
Dim j, p, x1, y1, colorSave, eraseIt As Boolean, pickColor As Boolean
Dim chkPix As Boolean, chgColor 'As Integer
Dim lineDraw As Boolean, lineOKDraw As Boolean
Dim rectDraw As Boolean, fillBoxDraw As Boolean, rectOKDraw As Boolean
Dim circleDraw As Boolean, fillCircleDraw As Boolean, circleOKDraw As Boolean
Dim textDraw As Boolean
Dim selRegion As Boolean
Dim lineX1, lineY1
Dim pasteIt As Boolean
'==For Select Region==================
Dim XHi, XLo, YHi, YLo, xDelLo, xdelHi, yDelLo, ydelHi
''Dim canMove As Boolean
Dim canSelect As Boolean
Dim moveIt As Boolean
Dim selectIt As Boolean
Dim xOff, yOff, setDiff As Boolean
Dim xMove, yMove

'========Used with Flood Area===================
Dim floodDraw As Boolean
Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Const FLOODFILLBORDER = 0
Const FLOODFILLSURFACE = 1
'===END Flood Area Data==================
'=======Used to Flip Image==========
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Private Sub cmdMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
MsgBox "Click on pixel in Expanded View", vbInformation
Exit Sub
End If
   ComDia.CancelError = True
   On Error GoTo ErrHandler
   ComDia.flags = cdlCCFullOpen
   ComDia.ShowColor
   If ComDia.Color = RGB(197, 197, 197) Then ComDia.Color = RGB(196, 196, 196)
If Button = 1 Then
PicMseColor.BackColor = ComDia.Color
Else
PicMseColorR.BackColor = ComDia.Color
End If
eraseIt = False
   Exit Sub
ErrHandler:
   ' User pressed Cancel button.
   Exit Sub

End Sub


Private Sub cmdRedo_Click()
DoReDo
picContainer.SetFocus
End Sub


Private Sub Form_Resize()
On Error Resume Next
'Move 0, 0, 9180, 7550
If Me.WindowState = 2 Then Me.WindowState = 0
Move FWhole.Width / 2 - Me.Width / 2, FWhole.Height / 2 - Me.Height / 2 - 400, 9180, 7550
End Sub

'Studio Image Start
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove_MouseDown 0, 0, 5, 5
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
B1_MouseMove 0, 0, 5, 5
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove_MouseUp 0, 0, 5, 5
End Sub

Private Sub B1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = True
End Sub

Private Sub Bmove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bover.Visible = True
End Sub

Private Sub Bmove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu FWhole.ms1, , B1.Left + B1.Width, B1.Top
Bmove.Visible = False
Bover.Visible = False
End Sub

'Studio Image End

Private Sub Form_Activate()
On Error Resume Next
Me.BackColor = RGB(63, 77, 158)
'Me.WindowState = 0
'Form_Resize
End Sub


Private Sub Form_Initialize()
On Error Resume Next
pic16Color.BackColor = RGB(63, 77, 158)
picBasic.BackColor = RGB(63, 77, 158)
Dim r, C, i, j
Dim ColorRect As RECT
  ''  picPaste.Top = 4
  ''  picPaste.Left = 4
cvtX = Screen.TwipsPerPixelX
cvtY = Screen.TwipsPerPixelY
''Clipboard.Clear


Toolbar1.Buttons(1).Caption = "新建"
Toolbar1.Buttons(2).Caption = "打开"
Toolbar1.Buttons(3).Caption = "保存"
Toolbar1.Buttons(5).Caption = "撤消"
Toolbar1.Buttons(6).Caption = "重复"
Toolbar1.Buttons(8).Caption = "剪切"
Toolbar1.Buttons(9).Caption = "复制"
Toolbar1.Buttons(10).Caption = "粘贴"
Toolbar1.Buttons(11).Caption = "删除"
Toolbar1.Buttons(12).Caption = "无筐选"
Toolbar1.Buttons(14).Caption = "辅助"

Toolbar1.Buttons(2).ButtonMenus(1).Text = "打开一个图标文件(.ico.cur.bmp.gif.jpg.wmf)"
Toolbar1.Buttons(2).ButtonMenus(2).Text = "加载任何一个文件的图标"

Label2.Caption = "图标预览"
Label6.Caption = "←点此查看16X16小图标"
Label5.Caption = "    此颜色       在保存后的图标中将成为透明色RGB(197,197,197)"
Label1.Caption = "当前画笔的颜色"
Label4(2).Caption = "更换画笔颜色，请点击下面"

Label4(0).Caption = "16种一般颜色"
Label4(1).Caption = "一些基本颜色"
cmdMore.Caption = "更多颜色"
cmdPicker.Caption = "从图中获取颜色"
Label3.Caption = "编辑区域 展开图"
Label8.Caption = "(点击更换)"

mnuFile.Caption = "文件(&F)"
mnuNew.Caption = "新建(&N)"
mnuOpen.Caption = "打开(&O)"
mnuSave.Caption = "保存(&S)"
mnuExit.Caption = "退出(&E)"

MnuEdit.Caption = "编辑(&E)"
MnuEditOpts(0).Caption = "撤消(&U)"
MnuEditOpts(1).Caption = "重复(&R)"
MnuEditOpts(3).Caption = "剪切(&X)"
MnuEditOpts(4).Caption = "复制(&C)"
MnuEditOpts(5).Caption = "粘贴(&V)"
MnuEditOpts(7).Caption = "删除(&D)"
MnuEditOpts(8).Caption = "清空剪贴板(&B)"
MnuEditOpts(9).Caption = "取消选择(&S)"

mnuOther.Caption = "工具(&T)"
mnuExtract.Caption = "加载任何一个文件的图标(&E)"
mnuAni.Caption = "建立动画光标(&A)"
nmuAni.Caption = "查看 .ani 文件(&V)"
mnuChgPix.Caption = "更改为透明图标(&C)"
mnuAssoc.Caption = ".ico 文件关联(&A)"

nmuEffects.Caption = "效果(&E)"
mnuHorz.Caption = "水平翻转(&H)"
mnuVert.Caption = "垂直翻转(&V)"
mnuRotateRight.Caption = "顺时针旋转90度(&R)"
mnuRotateLeft.Caption = "逆时针旋转90度(&L)"

mnuAbout2.Caption = "帮助(&H)"
mnuAbout.Caption = "关于"

Label10.Caption = "画室"
Label7.Caption = "效果"

mnuPopCut.Caption = "剪切(&X)"
mnuPopCopy.Caption = "复制(&C)"
mnuPopPaste.Caption = "粘贴(&V)"
mnuPopDelete.Caption = "删除(&D)"
mnuPopCancel.Caption = "取消选择(&S)"



Form1.Caption = "图标作坊 - 小画家"
chgColor = 0
picReal.ScaleMode = 3
picReal.ScaleHeight = 32
picReal.ScaleWidth = 32
picReal.Height = 32
picReal.Width = 32
'picReal.Top = 6
'picReal.Left = 211
picContainer.ScaleMode = 3
picContainer.ScaleHeight = 321
picContainer.ScaleWidth = 321
picContainer.Height = 321
picContainer.Width = 321
'picContainer.Top = 57
'picContainer.Left = 69
'Shape2.Left = 64
'Shape2.Top = 52
'Shape2.Width = 331
'Shape2.Height = 331
Line1.X2 = Form1.ScaleWidth
Set pic16Color.MouseIcon = picColorpicker.Picture
Set picBasic.MouseIcon = picColorpicker.Picture
'draw named colors
j = 0
For r = 1 To 2
        i = 1
    For C = 10 To 136 Step 18
        pic16Color.Line ((i * 18) - 8, ((r - 1) * 18) + 10)-((i * 18) - 8, ((r - 1) * 18) + 10), QBColor(j), BF
        i = i + 1
        j = j + 1
    Next C
Next r

'draw some basic colors
picBasic.DrawWidth = 16
picBasic.Line (10, 10)-(10, 10), RGB(255, 128, 128), BF
picBasic.Line (10, 28)-(10, 28), RGB(255, 0, 0), BF
picBasic.Line (28, 10)-(28, 10), RGB(255, 255, 128), BF
picBasic.Line (28, 28)-(28, 28), RGB(128, 255, 0), BF

picBasic.Line (46, 10)-(46, 10), RGB(0, 255, 128), BF
picBasic.Line (46, 28)-(46, 28), RGB(0, 255, 64), BF
picBasic.Line (64, 10)-(64, 10), RGB(128, 255, 255), BF
picBasic.Line (64, 28)-(64, 28), RGB(0, 255, 255), BF

picBasic.Line (82, 10)-(82, 10), RGB(0, 128, 255), BF
picBasic.Line (82, 28)-(82, 28), RGB(0, 128, 192), BF
picBasic.Line (100, 10)-(100, 10), RGB(255, 128, 192), BF
picBasic.Line (100, 28)-(100, 28), RGB(128, 128, 192), BF

picBasic.Line (118, 10)-(118, 10), RGB(255, 128, 255), BF
picBasic.Line (118, 28)-(118, 28), RGB(255, 0, 255), BF
picBasic.Line (136, 10)-(136, 10), RGB(128, 64, 64), BF
picBasic.Line (136, 28)-(136, 28), RGB(128, 0, 0), BF
'============
picBasic.Line (10, 46)-(10, 46), RGB(255, 128, 64), BF
picBasic.Line (10, 64)-(10, 64), RGB(255, 128, 0), BF
picBasic.Line (28, 46)-(28, 46), RGB(0, 255, 0), BF
picBasic.Line (28, 64)-(28, 64), RGB(0, 128, 0), BF

picBasic.Line (46, 46)-(46, 46), RGB(0, 128, 128), BF
picBasic.Line (46, 64)-(46, 64), RGB(0, 128, 128), BF
picBasic.Line (64, 46)-(64, 46), RGB(0, 64, 128), BF
picBasic.Line (64, 64)-(64, 64), RGB(0, 0, 255), BF

picBasic.Line (82, 46)-(82, 46), RGB(128, 128, 255), BF
picBasic.Line (82, 64)-(82, 64), RGB(0, 0, 160), BF
picBasic.Line (100, 46)-(100, 46), RGB(128, 0, 64), BF
picBasic.Line (100, 64)-(100, 64), RGB(128, 0, 128), BF

picBasic.Line (118, 46)-(118, 46), RGB(255, 0, 128), BF
picBasic.Line (118, 64)-(118, 64), RGB(128, 0, 255), BF
picBasic.Line (136, 46)-(136, 46), RGB(64, 0, 0), BF
picBasic.Line (136, 64)-(136, 64), RGB(0, 0, 0), BF
'=========
picBasic.Line (10, 82)-(10, 82), RGB(128, 64, 0), BF
picBasic.Line (10, 100)-(10, 100), RGB(128, 128, 0), BF
picBasic.Line (28, 82)-(28, 82), RGB(0, 64, 0), BF
picBasic.Line (28, 100)-(28, 100), RGB(128, 128, 64), BF

picBasic.Line (46, 82)-(46, 82), RGB(0, 64, 64), BF
picBasic.Line (46, 100)-(46, 100), RGB(128, 128, 128), BF
picBasic.Line (64, 82)-(64, 82), RGB(0, 0, 128), BF
picBasic.Line (64, 100)-(64, 100), RGB(64, 128, 128), BF

picBasic.Line (82, 82)-(82, 82), RGB(0, 0, 64), BF
picBasic.Line (82, 100)-(82, 100), RGB(64, 0, 64), BF
picBasic.Line (100, 82)-(100, 82), RGB(64, 0, 64), BF
picBasic.Line (100, 100)-(100, 100), RGB(64, 0, 128), BF

picBasic.Line (118, 82)-(118, 82), RGB(232, 144, 56), BF
picBasic.Line (118, 100)-(118, 100), RGB(255, 204, 153), BF
picBasic.Line (136, 82)-(136, 82), RGB(102, 51, 0), BF
picBasic.Line (136, 100)-(136, 100), RGB(18, 201, 55), BF
'With Form1.pic16Color
For r = 1 To 2
    For C = 0 To 143 Step 18
        SetRect ColorRect, (C + 2), (r * 18) - 16, C + 18, ((r - 1) * 18) + 18
        DrawEdge pic16Color.hdc, ColorRect, 2, 15
    Next C
Next r
'With Form1.picBasic
For r = 1 To 6
    For C = 0 To 143 Step 18
        SetRect ColorRect, (C + 2), (r * 18) - 16, C + 18, ((r - 1) * 18) + 18
        DrawEdge picBasic.hdc, ColorRect, 2, 15
    Next C
Next r
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim F, testExt, Answ
CheckSettings
PrepIconHeader
Form_Initialize

picTransparent.BackColor = RGB(197, 197, 197)
picReal.BackColor = RGB(197, 197, 197)
bkClr = RGB(197, 197, 197)
picContainer.BackColor = RGB(197, 197, 197) '&H8000000F
For F = 0 To picContainer.ScaleHeight Step 10
    picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
    picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
picReal.Picture = Image1.Picture
PaintDown
Form1.Refresh
UpdateUndo

'***Reg the file type****
Form_Resize
Me.Show

Dim GetV1, GetV1set, OpenPathV1 As String
 OpenPathV1 = App.Path & "\Studio.exe" & " /open %1"
GetV1 = getstring(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", "")
GetV1set = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "IcoLink")

If GetV1set = "" And GetV1 <> OpenPathV1 Then frmLink.Show 1

'*****Add Code to allow for Drag and Drop and Associations
If Command$ <> "" And Command$ <> "ap" And Command$ <> "ic" And Command$ <> "fp" Then
filePath = Command
If LCase(Left(Command, 6)) = "/open " Then
       filePath = StrConv(Mid(Command, 7), vbProperCase)
End If
filePath = LongFileName(filePath)
picReal.Picture = LoadPicture()
lblPath.Caption = filePath
testExt = Mid(filePath, Len(filePath) - 2, 3)
    If LCase(testExt) = "ico" Or LCase(testExt) = "jpg" Or LCase(testExt) = "gif" Or LCase(testExt) = "bmp" Or LCase(testExt) = "cur" Then
    '    MousePointer = 11
        picReal.BackColor = RGB(197, 197, 197)
        picTest.Picture = LoadPicture(filePath)
        DoEvents
            If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
                    ComDia.FileName = filePath
                    picReal.Picture = LoadPicture()
                    frmScroll.Show '1, Me
                    While Not iDone
                        DoEvents
                    Wend
                    iDone = False
            Else
                picReal = LoadPicture(filePath)
            End If
                PaintDown
                    cmdUndo.Visible = False
                    Toolbar1.Buttons(5).Enabled = False
        DeleteCollections
                UpdateUndo
                Form1.Refresh
                MousePointer = 0
'                lblPath.Caption = filePath
    Else
        ExtractRequest
    End If
End If
Form1.Show
cmdPaint_Click
Dirty = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
Bmove.Visible = False
Bover.Visible = False

Bmove2.Visible = False
Bover2.Visible = False

lblRGB.Caption = ""
lblMsePos.Caption = ""
'imgNew.Visible = True
'imgNewOver.Visible = False
'imgNewDown.Visible = False
'imgOpen.Visible = True
'imgOpenOver.Visible = False
'imgOpenDown.Visible = False
'imgSave.Visible = True
'imgSaveOver.Visible = False
'imgSaveDown.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply


If Dirty = True Then
    If LangA = "lgc" Then
reply = MsgBox("是否保存正在编辑的图片?", vbYesNoCancel + vbInformation, "File Not Saved")
Else
reply = MsgBox("Do you wish to save current image?", vbYesNoCancel + vbInformation, "File Not Saved")
End If
        
    If reply = vbCancel Then Cancel = 1 'Stop program from closing
    If reply = vbYes Then mnuSave_Click
End If
    If Dir(App.Path & "\temp.ico") <> "" Then Kill App.Path & "\temp.ico"
'==Be sure all forms are unloaded
Unload Form2
Unload Form3
Unload Form4
Unload frmAni
Unload ViewAni
'Unload MDIForm1
'=========================
If chkPix = True Then Cancel = 0
End Sub

Private Sub cmdPaint_Click()
On Error Resume Next
'=Initialize Switches=========
setSwitchesFalse
pixelDraw = True
'=====================
shSel.Left = cmdPaint.Left - 2
shSel.Top = cmdPaint.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdErase_Click()
'=Initialize Switches=========
setSwitchesFalse
eraseIt = True
pixelDraw = True
'=====================
shSel.Left = cmdErase.Left - 2
shSel.Top = cmdErase.Top - 2
picContainer.MouseIcon = picEraser.Picture
picContainer.SetFocus
End Sub

Private Sub cmdLine_Click()
'=Initialize Switches=========
setSwitchesFalse
lineDraw = True
'=====================
shSel.Left = cmdLine.Left - 2
shSel.Top = cmdLine.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdFlood_Click()
'=Initialize Switches=========
setSwitchesFalse
floodDraw = True
'=====================
shSel.Left = cmdFlood.Left - 2
shSel.Top = cmdFlood.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picFlood.Picture
End Sub

Private Sub cmdCircleDraw_Click()
'=Initialize Switches=========
setSwitchesFalse
circleDraw = True
'=====================
shSel.Left = cmdCircleDraw.Left - 2
shSel.Top = cmdCircleDraw.Top - 2
Shape1.Shape = 2 'Oval
Shape1.FillStyle = 1 'Transparent
picReal.FillStyle = 1 'transparent
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
picContainer.FillStyle = 1 'transparent
End Sub

Private Sub cmdFillCircleDraw_Click()
'=Initialize Switches=========
setSwitchesFalse
fillCircleDraw = True
'=====================
shSel.Left = cmdFillCircleDraw.Left - 2
shSel.Top = cmdFillCircleDraw.Top - 2
Shape1.Shape = 2 'Oval
Shape1.FillStyle = 0 'Solid
picReal.FillStyle = 0 'Solid
picContainer.FillStyle = 0 'Solid
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdFillBox_Click()
'=Initialize Switches=========
setSwitchesFalse
fillBoxDraw = True
'=====================
shSel.Left = cmdFillBox.Left - 2
shSel.Top = cmdFillBox.Top - 2
Shape1.Shape = 0 'Rectangle
Shape1.FillStyle = 0 'Solid
picReal.FillStyle = 0 'Solid
picContainer.FillStyle = 0 'Solid
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdRect_Click()
'=Initialize Switches=========
setSwitchesFalse
rectDraw = True
'=====================
shSel.Left = cmdRect.Left - 2
shSel.Top = cmdRect.Top - 2
Shape1.Shape = 0 'Rectangle
Shape1.FillStyle = 1 'Transparent
picReal.FillStyle = 1 'transparent
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
picContainer.FillStyle = 1 'transparent

End Sub

Private Sub cmdText_Click()
'=Initialize Switches=========
setSwitchesFalse
textDraw = True
'=====================
shSel.Left = cmdText.Left - 2
shSel.Top = cmdText.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picA.Picture
End Sub

Private Sub cmdRegion_Click()
setSwitchesFalse
picContainer.MouseIcon = picPencil.Picture
shSel.Left = cmdRegion.Left - 2
shSel.Top = cmdRegion.Top - 2
selRegion = True
Timer1.Enabled = True
moveIt = False
selectIt = True
picContainer.SetFocus
End Sub

Private Sub cmdUndo_Click()
DoUnDo
picContainer.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
FWhole.mIC.Enabled = True
FWhole.mIC.Checked = False
If FWhole.mAP.Checked = False Then
On Error Resume Next
 FWel.Show
 FWel.SetFocus
 Unload frmTray
  ' Load Tool List
End If
End Sub

Private Sub imgNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNew.Visible = False
imgNewOver.Visible = True
End Sub

Private Sub imgNewOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNewDown.Visible = True
End Sub
Private Sub imgNewOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpenOver.Visible = False
imgSaveOver.Visible = False
imgOpen.Visible = True
imgSave.Visible = True
End Sub

Private Sub imgNewOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Allows user to change mind.
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuNew_Click
imgNewDown.Visible = False
picContainer.SetFocus
End Sub

Private Sub imgOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpen.Visible = False
imgOpenOver.Visible = True
End Sub



Private Sub imgOpenOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpenDown.Visible = True
End Sub

Private Sub imgOpenOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpen.Visible = False
imgOpenOver.Visible = True
imgNewOver.Visible = False
imgSaveOver.Visible = False
imgSave.Visible = True
imgNew.Visible = True
End Sub

Private Sub imgOpenOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuOpen_Click
imgOpenDown.Visible = False
picContainer.SetFocus
End Sub



Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSave.Visible = False
imgSaveOver.Visible = True
End Sub




Private Sub imgSaveOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSaveDown.Visible = True
End Sub

Private Sub imgSaveOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSave.Visible = False
imgSaveOver.Visible = True
imgOpenOver.Visible = False
imgNewOver.Visible = False
imgOpen.Visible = True
imgNew.Visible = True
End Sub

Private Sub imgSaveOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuSave_Click
imgSaveDown.Visible = False
picContainer.SetFocus
End Sub




Private Sub Label6_Click()
picReal16.Picture = LoadPicture()
picReal.Picture = picReal.Image
picReal16.PaintPicture picReal.Picture, 0, 0, 16, 16
End Sub



'Studio Image Start
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove2_MouseDown 0, 0, 5, 5
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
B2_MouseMove 0, 0, 5, 5
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove2_MouseUp 0, 0, 5, 5
End Sub

Private Sub B2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove2.Visible = True
End Sub

Private Sub Bmove2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bover2.Visible = True
End Sub

Private Sub Bmove2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu Form1.nmuEffects, , b2.Left + b2.Width, b2.Top
Bmove2.Visible = False
Bover2.Visible = False
End Sub


'Studio Image End



Private Sub mnuAni_Click()
Form1.Hide
frmAni.Show
End Sub

Private Sub mnuAssoc_Click()

frmLink.Show 1
'Dim Ret
'Ret = MsgBox("This option will associate Window Icon (.ico) files with this program so that double clicking any ico file will open it in this program. Is that what you want to do?", vbYesNo + vbInformation, "Associate icon files")
'If Ret = vbNo Then Exit Sub
'SetUpIconDblClick
End Sub

Private Sub mnuChgPix_Click()
Dim X, Y, ret
lineDraw = False
If LangA = "lgc" Then
ret = MsgBox("请在编辑区域中点击一个需要成为透明颜色的像素点", vbOKCancel + vbInformation, "Change Pixel Color")
Else
ret = MsgBox("Click On Pixel Color In Expanded View To Be Made Transparent. ", vbOKCancel + vbInformation, "Change Pixel Color")
End If
If ret = vbCancel Then Exit Sub
chkPix = True
'=Initialize Switches=========
setSwitchesFalse
'=====================
While chgColor = 0
DoEvents
Wend
'MousePointer = 11
picReal.Picture = picReal.Image
For X = 0 To 31
For Y = 0 To 31
If picReal.Point(X, Y) = chgColor Then
picReal.PSet (X, Y), RGB(197, 197, 197)
End If
Next Y
Next X
picReal.Picture = picReal.Image
PaintDown
Form1.Refresh
cmdPaint_Click
MousePointer = 0
chgColor = 0
chkPix = False
Dirty = True
End Sub

Private Sub mnuDelete_Click()
moveIt = True
selectIt = False
End Sub

Private Sub MnuEdit_Click()
    MnuEditOpts(0).Enabled = ColUndo.Count > 1
    MnuEditOpts(1).Enabled = ColRedo.Count > 0
    MnuEditOpts(3).Enabled = shRect.Visible
     MnuEditOpts(4).Enabled = shRect.Visible
    MnuEditOpts(5).Enabled = Clipboard.GetFormat(vbCFBitmap) 'shRect.Visible
    MnuEditOpts(7).Enabled = shRect.Visible
    MnuEditOpts(8).Enabled = Clipboard.GetFormat(vbCFBitmap) 'shRect.Visible
    MnuEditOpts(9).Enabled = shRect.Visible
End Sub

Private Sub MnuEditOpts_Click(Idx%)

    Select Case Idx
           Case 0
                DoUnDo
           Case 1
                DoReDo
           Case 3
                mnuPopCut_Click
           Case 4
                mnuPopCopy_Click
           Case 5
                pasteItNow
           Case 7
                mnuPopDelete_Click
           Case 8
                Clipboard.Clear
           Case 9
                cmdRegion_Click
    End Select
End Sub

Private Sub mnuHelp_Click()
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\icon.htm", "", App.Path, 1
End Sub

Private Sub mnuNew_Click()
chkSave
Dim F

'=======ClearUndo
    UpdateUndo
    cmdUndo.Visible = False
    Toolbar1.Buttons(5).Enabled = False
        DeleteCollections
'=Initialize Switches=========
setSwitchesFalse
'=====================
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
picContainer.Picture = LoadPicture()
picContainer.Cls
PaintDown
picReal.BackColor = RGB(197, 197, 197)
picContainer.BackColor = RGB(197, 197, 197)
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
lblPath.Caption = "Untitled"
Form1.Refresh
cmdPaint_Click
UpdateUndo
Dirty = False
End Sub

Private Sub mnuOpen_Click()
Dim Answ, cdDir, cdIndex, Pos
imgOpen.Visible = True
imgOpenOver.Visible = False
imgOpenDown.Visible = False
chkSave
'=ClearUndo========
    UpdateUndo
    cmdUndo.Visible = False
    Toolbar1.Buttons(5).Enabled = False
    DeleteCollections
'=Initialize Switches=========
setSwitchesFalse
'=====================
ComDia.CancelError = True
On Error GoTo ex
ComDia.FileName = ""
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdOpenDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdOpenIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
NoRegVal: ComDia.flags = cdlOFNFileMustExist
ComDia.Filter = "Icons (*.ico;*.cur)|*.ico;*.cur|Images (*.bmp;*.jpg;*gif;*wmf)|*.bmp;*.jpg;*.gif;*wmf"
ComDia.ShowOpen
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdOpenDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdOpenIndexSetting", cdIndex
'MousePointer = 11
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
    frmScroll.Show '1, Me
    While Not iDone
    DoEvents
    Wend
    iDone = False
Else
    picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
Form1.Refresh
MousePointer = 0
lblPath.Caption = ComDia.FileName
cmdPaint_Click
UpdateUndo
Dirty = False

Exit Sub
ex: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description, vbInformation, "Error"
Exit Sub
End Sub

Private Sub mnuExtract_Click()
Dim Answ, cdDir, cdIndex, Pos
   Dim hImgLarge As Long
   Dim hImgSmall As Long   'the handle to the system image list
   Dim fName As String     'the file name to get icon from
   Dim fnFilter As String  'the file name filter
   Dim r As Long
chkSave
Dirty = False
'ClearUndo
    UpdateUndo
    DeleteCollections
    
'=Initialize Switches=========
setSwitchesFalse
'=====================
   
   On Local Error GoTo cmdLoadErrorHandler
   
  'get the file from the user
   fnFilter$ = "All Files (*.*)|*.*"
'==========
ComDia.FileName = ""
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdExtractDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdExtractIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
NoRegVal: ComDia.flags = cdlOFNFileMustExist
'===========

   ComDia.CancelError = True
   ComDia.Filter = fnFilter$
   ComDia.ShowOpen
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdExtractDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdExtractIndexSetting", cdIndex
'============
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
   fName$ = ComDia.FileName
   
'get the system icon associated with that file
   hImgSmall& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

   hImgLarge& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   
   picTest.Picture = LoadPicture()
   picTest.AutoRedraw = True

   picTest.BackColor = RGB(197, 197, 197)
  'draw the associated icon into the pictureboxes
   Call ImageList_Draw(hImgLarge&, shinfo.iIcon, picReal.hdc, 0, 0, ILD_TRANSPARENT)

PaintDown
Form1.Refresh
UpdateUndo
cmdUndo.Visible = False
Toolbar1.Buttons(5).Enabled = False
MousePointer = 0
lblPath.Caption = ComDia.FileName
cmdPaint_Click
Dirty = False
Exit Sub

cmdLoadErrorHandler: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description, vbInformation, "Error"
Exit Sub
End Sub



Private Sub mnuPopCancel_Click()
cmdRegion_Click
End Sub

Private Sub mnuPopCopy_Click()
Clipboard.Clear
Clipboard.SetData picClip.Clip
            '===========Allow another Select Region to be made========
            cmdRegion_Click
UpdateUndo
End Sub

Private Sub mnuPopCut_Click()

Dim C
Clipboard.Clear
Clipboard.SetData picClip.Clip
               For r = yDelLo To ydelHi
                  For C = xDelLo To xdelHi
                       picReal.PSet (C, r), RGB(197, 197, 197)
                   Next C
               Next r
        PaintDown
        UpdateUndo
        Form1.Refresh
        cmdRegion_Click
End Sub

Private Sub mnuPopDelete_Click()
Dim C
Dirty = True
               For r = yDelLo To ydelHi
                  For C = xDelLo To xdelHi
                       picReal.PSet (C, r), RGB(197, 197, 197)
                   Next C
               Next r
        PaintDown
        UpdateUndo
        Form1.Refresh
        cmdRegion_Click
End Sub

Private Sub pasteItNow()
Dim r%, C%
picContainer.Picture = picContainer.Image
Dirty = True
cmdRegion_Click
okToMove = True
pasteIt = True
shRect.Visible = False
picMove.Visible = False
If Clipboard.GetFormat(vbCFBitmap) Then
        picCB.Picture = Clipboard.GetData(vbCFBitmap)
    If picCB.Width > 32 Or picCB.Height > 32 Then
        IsCB = True
        frmScroll.Show '1, Me
        While Not iDone
        DoEvents
        Wend
        iDone = False
    End If
        DoEvents
        picMove.Width = 10 * picCB.Width
        picMove.Height = 10 * picCB.Height
        picCB.Picture = picCB.Image
        '==build mask and sprite
        picCBmask.Width = picCB.Width
        picCBmask.Height = picCB.Height
        picCBsprite.Width = picCB.Width
        picCBsprite.Height = picCB.Height
        For r = 0 To picCB.Height - 1
            For C = 0 To picCB.Width - 1
                If picCB.Point(C, r) <> RGB(197, 197, 197) Then
                    picCBmask.PSet (C, r), vbBlack
                    picCBsprite.PSet (C, r), picCB.Point(C, r)
                Else
                    picCBmask.PSet (C, r), vbWhite
                    picCBsprite.PSet (C, r), vbBlack
                End If
            Next C
        Next r
            picCBmask.Picture = picCBmask.Image
            picCBsprite.Picture = picCBsprite.Image
        '========

End If
        shRect.Move 0, 0, picMove.Width + 4, picMove.Height + 4
        picMove.Move 4, 4
        shRect.Visible = True
        'paint to picContainer image
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBmask.Width * 10, picCBmask.Height * 10, picCBmask.hdc, 0, 0, picCBmask.Width, picCBmask.Height, vbSrcAnd
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBsprite.Width * 10, picCBsprite.Height * 10, picCBsprite.hdc, 0, 0, picCBsprite.Width, picCBsprite.Height, vbSrcPaint
        DoEvents
        
End Sub

Private Sub mnuRotateLeft_Click()
    Picture2.Cls
    Call MovePixelsLeft
    picReal.Picture = LoadPicture()
    Picture2.Picture = Picture2.Image
    picReal.Picture = Picture2.Picture
    picReal.Refresh
    UpdateUndo
    PaintDown
    Form1.Refresh
    Dirty = True
End Sub
Private Sub MovePixelsLeft()
Dim r, C, p
For C = 0 To 31
    For r = 0 To 31
        p = picReal.Point(C, r)
        Picture2.PSet (r, 31 - C), p
    Next r
Next C
End Sub
Private Sub mnuRotateRight_Click()
    Picture2.Cls
    Call MovePixelsRight
    picReal.Picture = LoadPicture()
    Picture2.Picture = Picture2.Image
    picReal.Picture = Picture2.Picture
    picReal.Refresh
    UpdateUndo
    PaintDown
    Form1.Refresh
    Dirty = True
End Sub
Private Sub MovePixelsRight()
Dim r, C, p
For r = 0 To 31
    For C = 0 To 31
        p = picReal.Point(C, r)
        Picture2.PSet (31 - r, C), p
    Next C
Next r
End Sub


Private Sub mnuSave_Click()
Dim ret, bmpPicInfo As BITMAPINFO, Answ, cdDir, cdIndex, Pos
Dim sPos, ePos
imgSave.Visible = True
imgSaveOver.Visible = False
imgSaveDown.Visible = False
'==========
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdSaveAsDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdSaveAsIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
'===========
NoRegVal: ComDia.CancelError = True
On Error GoTo ExitIt

ComDia.FileName = "Created"
ePos = InStrRev(lblPath.Caption, ".")
sPos = InStrRev(lblPath.Caption, "\")
If ePos > 0 Then
ComDia.FileName = LCase(Mid(lblPath.Caption, sPos + 1, (ePos - sPos) - 1))
End If
ComDia.flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
ComDia.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp|Cursors (*.cur)|*.cur"
ComDia.ShowSave
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdSaveAsDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdSaveAsIndexSetting", cdIndex
'============

If ComDia.FilterIndex = 1 Then 'Save as Icon
   ' MousePointer = 11
    Form2.Show 1
    If CancelIt = True Then 'User pressed Cancel on SaveAsOptions Form
    CancelIt = False
    MousePointer = 0
    Exit Sub
    End If
    '============
    lblPath.Caption = ComDia.FileName
    If Form2.opt1Bit Then BitCnt = 1
    If Form2.opt4Bit Then BitCnt = 4
    If Form2.opt8Bit Then BitCnt = 8
    If Form2.opt24Bit Then BitCnt = 24
    '=================
        With bmpPicInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bmpPicInfo.bmiHeader)
        .biWidth = 32
        .biHeight = 32
    End With
    IconInfo.iDC = CreateCompatibleDC(0)
    IconInfo.iWidth = 32
    IconInfo.iHeight = 32
    bi24BitInfo.bmiHeader.biWidth = 32
    bi24BitInfo.bmiHeader.biHeight = 32
    IconInfo.iBitmap = CreateDIBSection(IconInfo.iDC, bmpPicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject IconInfo.iDC, IconInfo.iBitmap
    ret = BitBlt(IconInfo.iDC, 0, 0, 32, 32, picReal.hdc, 0, 0, vbSrcCopy)
    If ret = 0 Then
    MsgBox "Unable to BitBlt Picture.", vbCritical, "Error"
    Exit Sub
    End If
    DoEvents
    SaveIcon ComDia.FileName, IconInfo.iDC, IconInfo.iBitmap, BitCnt ', CLng(SaveTypeIn)
    IconInfo.iFileName = ComDia.FileName
    DeleteDC IconInfo.iDC
    DeleteObject IconInfo.iBitmap
    '==================
    picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then

If LangA = "lgc" Then
Answ = MsgBox("图象不是 32X32 像素格式. 是否调整其大小?", vbYesNo + vbInformation, "Image Size Test")
Else
Answ = MsgBox("Image not in 32X32 pixel format. Do you wish to resize?", vbYesNo + vbInformation, "Image Size Test")
End If



    If Answ = vbYes Then
    picReal.PaintPicture picTest.Image, 0, 0, 32, 32
    Else
        Exit Sub
    End If
Else
picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
Form1.Refresh
'======new method to get rid of black ===========
If BitCnt = 24 Then
    Dim hIcon
PicIcon.Picture = LoadPicture()
PicIcon.Cls
    ExtractIconEx ComDia.FileName, 0, hIcon, 0, 1
    ret = DrawIconEx(PicIcon.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H3&) 'Const DI_NORMAL = &H3 Both Mask and Image
    If ret = 0 Then
        MsgBox "Unable to draw PicIcon", vbInformation
    End If
    PicIcon.Refresh
    ret = DrawIconEx(picMask.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H1&) 'Const DI_MASK = &H1
    If ret = 0 Then
        MsgBox "Unable to draw picMask", vbInformation
    End If
    picMask.Refresh
    ret = DrawIconEx(PicImage.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H2&) 'Const DI_IMAGE = &H2
    If ret = 0 Then
        MsgBox "Unable to draw PicImage", vbInformation
    End If
    PicImage.Refresh
    DestroyIcon hIcon
    
    WriteDataToFile ComDia.FileName

End If
    '==================
    MousePointer = 0
    
End If
'=========================
If ComDia.FilterIndex = 2 Then 'Save as Bmp
    SavePicture picReal.Image, ComDia.FileName
    lblPath.Caption = ComDia.FileName
End If
'========================
If ComDia.FilterIndex = 3 Then 'Save as Cur
Answ = MsgBox("Do you wish to save in Color to be used in creating an animated (.ani) cursor file later?", vbYesNo + vbInformation, "Color or B-W?")
  '  MousePointer = 11
    If Answ = vbNo Then
        BitCnt = 1
    Else
        BitCnt = 24
    End If
    '=================
        With bmpPicInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bmpPicInfo.bmiHeader)
        .biWidth = 32
        .biHeight = 32
    End With
    IconInfo.iDC = CreateCompatibleDC(0)
    IconInfo.iWidth = 32
    IconInfo.iHeight = 32
    bi24BitInfo.bmiHeader.biWidth = 32
    bi24BitInfo.bmiHeader.biHeight = 32
    IconInfo.iBitmap = CreateDIBSection(IconInfo.iDC, bmpPicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject IconInfo.iDC, IconInfo.iBitmap
    ret = BitBlt(IconInfo.iDC, 0, 0, 32, 32, picReal.hdc, 0, 0, vbSrcCopy)
    If ret = 0 Then
    MsgBox "Unable to BitBlt Picture.", vbInformation
    Exit Sub
    End If
    DoEvents
    If Dir(App.Path & "\temp.ico") <> "" Then Kill App.Path & "\temp.ico"
    SaveIcon App.Path & "\temp.ico", IconInfo.iDC, IconInfo.iBitmap, BitCnt ', CLng(SaveTypeIn)
    DeleteDC IconInfo.iDC
    DeleteObject IconInfo.iBitmap
    DoEvents
    '==================

 Form3.Show 1
    If CancelIt = True Then 'User pressed Cancel on SetCursorHotspots Form
    CancelIt = False
    MousePointer = 0
    Exit Sub
    End If
    lblPath.Caption = ComDia.FileName
    '==========
    
    picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
If LangA = "lgc" Then
Answ = MsgBox("图象不是 32X32 像素格式. 是否调整其大小?", vbYesNo + vbInformation, "Image Size Test")
Else
Answ = MsgBox("Image not in 32X32 pixel format. Do you wish to resize?", vbYesNo + vbInformation, "Image Size Test")
End If
    If Answ = vbYes Then
    picReal.PaintPicture picTest.Image, 0, 0, 32, 32
    Else
        Exit Sub
    End If
Else
picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
MousePointer = 0
    '==================
    MousePointer = 0
End If

Dirty = False
Form1.Refresh
Exit Sub
ExitIt: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description, vbInformation
End Sub
Private Sub WriteDataToFile(Fn$)

    Dim MaskString$
    Dim mSG$
    Dim F%, H%, W%
    Dim c1&, c2&, r&, G&, B&, k%, n%

    On Error GoTo WriteError

    F = FreeFile
    Open Fn For Binary Access Write As #F


         For k = Len(Fn) To 1 Step -1
             If Mid(Fn, k, 1) = "\" Then Exit For
         Next

         Put #F, 1, ID
         Put #F, 7, IDE
         Put #F, 23, BIH
         k = 63
         For H = 31 To 0 Step -1
             For W = 0 To 31
                 c1 = GetPixel(PicImage.hdc, W, H)
                 c2 = GetPixel(picMask.hdc, W, H)
                 If c2 = &HFFFFFF Then
                    Put #F, k, 0
                    Put #F, k + 1, 0
                    Put #F, k + 2, 0
                 Else
                    B = c1 \ 65536
                    G = (c1 - B * 65536) \ 256
                    r = c1 - B * 65536 - G * 256
                    Put #F, k, B
                    Put #F, k + 1, G
                    Put #F, k + 2, r
                 End If
                 k = k + 3
             Next
         Next
         k = 0
         n = 0
         For H = 31 To 0 Step -1
             For W = 0 To 31
                 If GetPixel(picMask.hdc, W, H) = &HFFFFFF Then
                    MaskString = MaskString & "1"
                 Else
                    MaskString = MaskString & "0"
                 End If
                 k = k + 1
                 If k = 8 Then
                    k = 0
                    Put #F, n + 3135, BinaryStringToByte(MaskString)
                    MaskString = ""
                    n = n + 1
                 End If
             Next
         Next
    Close #F

    CurrentFile = Fn
    On Error GoTo 0
    Exit Sub

WriteError:

    Screen.MousePointer = 0

    If Err.Number <> cdlCancel Then
       mSG = Err.Description & "."
       mSG = mSG & vbCrLf & vbCrLf
       If CurrentFile = "Untitled" Then
          mSG = mSG & "Unable to save Untitled."
       Else
          mSG = mSG & "Unable to save " & CurrentName
       End If
       MsgBox mSG, vbExclamation, Ttl & " - Error"
    End If
    'bFileSaved = False
    Err.Clear
    Exit Sub

End Sub
Private Function BinaryStringToByte(MS$) As Byte

    Dim k%, Rv As Byte

    For k = 1 To 8
        If Mid(MS, k, 1) = "1" Then Rv = Rv + 2 ^ (8 - k)
    Next

    BinaryStringToByte = Rv

End Function
Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub cmdPicker_Click()
If selRegion Then Exit Sub
'If chkPix = True Then
    If LangA = "lgc" Then
    MsgBox "现在请点击编辑区域的像素点"
    Else
    MsgBox "Now click on pixel in Expanded View"
    End If

'Exit Sub
'End If
pickColor = True
picContainer.MouseIcon = picColorpicker.Picture
End Sub

Private Sub mnuAbout_Click()
about1.Show


End Sub

Public Sub PaintDown()
Dim F
Static Pont
Pont = picContainer.Point(0, 0)
picContainer.PaintPicture picReal.Image, 0, 0, 321, 321
If Pont = &HFFC0FF Then
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &HFFC0FF
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &HFFC0FF
Next F
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(15), B
Else
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(0), B
End If
End Sub

Private Sub mnuHorz_Click()
Dim pX As Long, pY As Long, retval As Long
On Error GoTo errMsg
picTemp.Cls
pX = picReal.ScaleWidth
pY = picReal.ScaleHeight
picTemp.Width = picReal.Width
picTemp.Height = picReal.Height
retval = StretchBlt(picTemp.hdc, pX - 1, 0, -pX, pY, _
picReal.hdc, 0, 0, pX, pY, SRCCOPY)
picReal.Cls
picTemp.Picture = picTemp.Image
picReal.PaintPicture picTemp.Picture, 0, 0, _
picTemp.Width, picTemp.Height, 0, 0, _
picTemp.Width, picTemp.Height, vbSrcCopy
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
Exit Sub
errMsg: MsgBox "Error # " & Err.Number & " " & Err.Description, vbInformation
Err.Clear
picTemp.Cls
picTemp.Picture = LoadPicture()
End Sub

Private Sub mnuVert_Click()
Dim pX As Long, pY As Long, retval As Long
On Error GoTo errMsg
picTemp.Cls
pX = picReal.ScaleWidth
pY = picReal.ScaleHeight
picTemp.Width = picReal.Width
picTemp.Height = picReal.Height
retval = StretchBlt(picTemp.hdc, 0, pY - 1, pX, -pY, _
picReal.hdc, 0, 0, pX, pY, SRCCOPY)
picReal.Cls
picTemp.Picture = picTemp.Image
picReal.PaintPicture picTemp.Picture, 0, 0, _
picTemp.Width, picTemp.Height, 0, 0, _
picTemp.Width, picTemp.Height, vbSrcCopy
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
Exit Sub
errMsg: MsgBox "Error # " & Err.Number & " " & Err.Description, vbInformation
Err.Clear
picTemp.Cls
picTemp.Picture = LoadPicture()

End Sub

Private Sub nmuAni_Click()
Form1.Hide
ViewAni.Show
End Sub

Private Sub pic16Color_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
If LangA = "lgc" Then
MsgBox "请点击编辑区域的像素点", vbInformation
Else
MsgBox "Click on pixel in Expanded View", vbInformation
End If
Exit Sub
End If
If Button = 1 Then
PicMseColor.BackColor = pic16Color.Point(X, Y)
Else
PicMseColorR.BackColor = pic16Color.Point(X, Y)
End If
End Sub

Private Sub pic16Color_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = pic16Color.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr
End Sub

Private Sub picBasic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
If LangA = "lgc" Then
MsgBox "请点击编辑区域的像素点", vbInformation
Else
MsgBox "Click on pixel in Expanded View", vbInformation
End If
Exit Sub
End If
If Button = 1 Then
PicMseColor.BackColor = picBasic.Point(X, Y)
Else
PicMseColorR.BackColor = picBasic.Point(X, Y)
End If
End Sub

Private Sub picBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = picBasic.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr

End Sub



Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cr
'========Move Paste Image===
'In picContainer Mouse down, move and up -- 'Move Paste Image'
'must be before 'Select Region'
If okToMove Then 'set True in PasteItNow
    moveIt = True
    pX = X
    pY = Y
    prevX = X \ 10
    prevY = Y \ 10
    pXOff = X - shRect.Left
    pYOff = Y - shRect.Top
    Exit Sub
End If
'=========Select Region==========
If selRegion = True Then
    If selectIt Then
        canSelect = True
        xStart = (X \ 10) * 10
        yStart = (Y \ 10) * 10
        XLo = (X \ 10) * 10
        YLo = (Y \ 10) * 10
        XHi = (X \ 10) * 10
        YHi = (Y \ 10) * 10
        shRect.Width = Abs(XHi - XLo)
        shRect.Height = Abs(YHi - YLo)
        Exit Sub
    End If
End If
'=========Pick a Color======================
If pickColor = True Then
If Button = 1 Then
PicMseColor.BackColor = picContainer.Point(X, Y)
Else
PicMseColorR.BackColor = picContainer.Point(X, Y)
End If
picContainer.MouseIcon = picPencil.Picture
    eraseIt = False
    pickColor = False
    Exit Sub
End If
'========Draw Text=================
If textDraw Then
Dirty = True
        If Button = 1 Then
            picReal.ForeColor = PicMseColor.BackColor
            Form4.Text1.ForeColor = PicMseColor.BackColor
        Else
            picReal.ForeColor = PicMseColorR.BackColor
            Form4.Text1.ForeColor = PicMseColorR.BackColor
        End If
curX = X \ 10
curY = Y \ 10
DoEvents
Form4.Show 1
        picReal.Picture = picReal.Image
        UpdateUndo
        PaintDown
        Form1.Refresh
End If
'=========Flood an Area===========
If floodDraw Then
Dirty = True
picReal.FillStyle = 0 'Solid
        If Button = 1 Then
            picReal.FillColor = PicMseColor.BackColor
        Else
            picReal.FillColor = PicMseColorR.BackColor
        End If
ExtFloodFill picReal.hdc, X \ 10, Y \ 10, picReal.Point(X \ 10, Y \ 10), FLOODFILLSURFACE
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
End If
'=========Draw a Line=======================
If lineDraw Then
    lineOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
'=========Draw a Rectangle Or FillBox=======================
If rectDraw Then
    rectOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
If fillBoxDraw Then
    rectOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
'=========Draw a Circle or Filled Circle=======================
If circleDraw Then
    circleOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
If fillCircleDraw Then
    circleOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If

'=========Change a Pixel Color to Transparent=======
If chkPix = True Then
chgColor = picContainer.Point(X, Y)
chkPix = False
Exit Sub
End If
'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
canDraw = True
End If
End Sub

Sub ReEnToolbar()
Toolbar1.Buttons(8).Enabled = shRect.Visible
Toolbar1.Buttons(9).Enabled = shRect.Visible
Toolbar1.Buttons(10).Enabled = Clipboard.GetFormat(vbCFBitmap) 'shRect.Visible
Toolbar1.Buttons(11).Enabled = shRect.Visible
Toolbar1.Buttons(12).Enabled = shRect.Visible
End Sub


Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReEnToolbar
Dim cr, xM, yM
'==========Show Mouse Position=======
xM = Int(X / 10) + 1
yM = Int(Y / 10) + 1
lblMsePos.Caption = "Mouse (X,Y) is " & xM & "," & yM
'=========Show color of pixel mouse is over========
clr = picContainer.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr
'===========Show Hand on Paste image========
        If okToMove = True And X > shRect.Left And X < shRect.Left + shRect.Width And Y > shRect.Top And Y < shRect.Top + shRect.Height Then
            picContainer.MouseIcon = picHand.Picture
        Else
            picContainer.MouseIcon = picPencil.Picture
        End If
'=========Move Paste image=============
    If moveIt Then
        If X \ 10 = prevX And Y \ 10 = prevY Then
            prevX = X \ 10
            prevY = Y \ 10
            Exit Sub
        Else
            Timer1.Enabled = False
            shRect.Visible = False
            picContainer.Cls
            shRect.Left = X - pXOff
            shRect.Top = Y - pYOff
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBmask.Width * 10, picCBmask.Height * 10, picCBmask.hdc, 0, 0, picCBmask.Width, picCBmask.Height, vbSrcAnd
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBsprite.Width * 10, picCBsprite.Height * 10, picCBsprite.hdc, 0, 0, picCBsprite.Width, picCBsprite.Height, vbSrcPaint
            prevX = X \ 10
            prevY = Y \ 10
          Exit Sub
        End If
    End If


'========Select Region==========
If selRegion Then

    If X < xStart And Y < yStart Then
            XLo = xStart + 10
            YLo = yStart + 10
            XHi = ((X \ 10) * 10)
            YHi = ((Y \ 10) * 10)
            
    End If
    If X > xStart And Y > yStart Then
            XLo = xStart
            YLo = yStart
            XHi = ((X \ 10) * 10) + 10
            YHi = ((Y \ 10) * 10) + 10
    End If
    If X > xStart And Y < yStart Then
            XLo = xStart
            YLo = yStart + 10
            XHi = ((X \ 10) * 10) + 10
            YHi = ((Y \ 10) * 10)
    End If
    If X < xStart And Y > yStart Then
            XLo = xStart + 10
            YLo = yStart
            XHi = ((X \ 10) * 10)
            YHi = ((Y \ 10) * 10) + 10
    End If
    If XHi < 0 Then XHi = 0
    If YHi < 0 Then YHi = 0
    If XHi > picContainer.ScaleWidth - 1 Then XHi = picContainer.ScaleWidth - 1
    If YHi > picContainer.ScaleHeight - 1 Then YHi = picContainer.ScaleHeight - 1
    If XLo < 0 Then XLo = 0
    If YLo < 0 Then YLo = 0
    If XLo > picContainer.ScaleWidth - 1 Then XLo = picContainer.ScaleWidth - 1
    If YLo > picContainer.ScaleHeight - 1 Then YLo = picContainer.ScaleHeight - 1
    If canSelect = True Then
            shRect.Width = Abs(XHi - XLo)
            shRect.Height = Abs(YHi - YLo)
            shRect.Visible = True
        If XHi > XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XLo
        End If
        If XHi > XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XLo
        End If
        If XHi < XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XHi
        End If
        If XHi < XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XHi
        End If
    End If
Exit Sub
End If
'=======Pick A Color and Flood=================
If pickColor Then Exit Sub
If floodDraw Then Exit Sub
'=========Draw a Line=======================
    If lineOKDraw Then
        Line2.x1 = lineX1
        Line2.y1 = lineY1
        Line2.X2 = X
        Line2.Y2 = Y
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
        Line2.BorderColor = cr
        Line2.Visible = True

        Exit Sub
    End If
'=========Draw a Rectangle or Fill Box=======================
    If rectDraw Or fillBoxDraw Then
        If rectOKDraw = True Then
                If X > lineX1 Then
                    Shape1.Left = lineX1
                Else
                    Shape1.Left = X
                End If
                
                If Y > lineY1 Then
                    Shape1.Top = lineY1
                Else
                    Shape1.Top = Y
                End If
                Shape1.Width = Abs(X - lineX1)
                Shape1.Height = Abs(Y - lineY1)
                
                If Button = 1 Then
                    cr = PicMseColor.BackColor
                Else
                    cr = PicMseColorR.BackColor
                End If
            Shape1.BorderColor = cr
            Shape1.FillColor = cr
            Shape1.Visible = True
            Exit Sub
        End If
    End If
'=========Draw a Circle or Filled Circle=======================
    If circleDraw Or fillCircleDraw Then
        If circleOKDraw = True Then
                Shape1.Width = Abs(X - lineX1)
                Shape1.Height = Abs(Y - lineY1)
        Shape1.Visible = True
            If X > lineX1 Then
                    Shape1.Left = lineX1 - (Shape1.Width / 2)
                Else
                    Shape1.Left = X + (Shape1.Width / 2)
                End If
                
                If Y > lineY1 Then
                    Shape1.Top = lineY1 - (Shape1.Height / 2)
                Else
                    Shape1.Top = Y + (Shape1.Height / 2)
            End If

                
                If Button = 1 Then
                    cr = PicMseColor.BackColor
                Else
                    cr = PicMseColorR.BackColor
                End If
            Shape1.BorderColor = cr
            Shape1.FillColor = cr
            Shape1.Visible = True
            Exit Sub
        End If
    End If
'=========Change a Pixel Color to Transparent=======

If chkPix = True Then Exit Sub

'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
    If canDraw = True Then
                Dirty = True
            If Button = 1 Then
                ColorChg = PicMseColor.BackColor
            Else
                ColorChg = PicMseColorR.BackColor
            End If
                If eraseIt Then ColorChg = RGB(197, 197, 197) 'Transparent picContainer.BackColor
            x1 = 0
            y1 = 0
        For j = 0 To 31
        For p = 0 To 31

                If X < x1 + 10 And X > x1 And Y < y1 + 10 And Y > y1 Then
                    picContainer.Line (x1 + 1, y1 + 1)-(x1 + 9, y1 + 9), ColorChg, BF
                    picReal.PSet (x1 \ 10, y1 \ 10), ColorChg
                End If

            x1 = x1 + 10
                If x1 = 320 Then
                    x1 = 0
                    y1 = y1 + 10
                End If


        Next p
        Next j
    End If
End If

End Sub

Private Sub picContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cr, F, xd, yd, r, C
On Error GoTo ExitIt
If pickColor Then Exit Sub
If floodDraw Then Exit Sub
'=======Move Paste Image===========
If moveIt = True Then
'=====paint non transparent to picReal from picCB
xd = (shRect.Left + 2) \ 10
yd = (shRect.Top + 2) \ 10
BitBlt picReal.hdc, xd, yd, picCBmask.Width, picCBmask.Height, picCBmask.hdc, 0, 0, vbSrcAnd
BitBlt picReal.hdc, xd, yd, picCBsprite.Width, picCBsprite.Height, picCBsprite.hdc, 0, 0, vbSrcPaint
picCB.Picture = LoadPicture()
picCBmask.Picture = LoadPicture()
picCBsprite.Picture = LoadPicture()
            picReal.Picture = picReal.Image
            UpdateUndo
            PaintDown
            Form1.Refresh
    cmdRegion_Click
    moveIt = False
    Exit Sub
End If
'=======Select Region==============
If selRegion Then

    If canSelect Then
            'Clip Screen Image
            picReal.Picture = picReal.Image
            picClip.Picture = picReal.Picture
            picContainer.Picture = picContainer.Image
            PicClipMove.Picture = picContainer.Picture
            DoEvents
            ' Get X and Y coordinates of the clipping region.
            picClip.ClipX = (shRect.Left \ 10)
            PicClipMove.ClipX = shRect.Left
            picClip.ClipY = (shRect.Top \ 10)
            PicClipMove.ClipY = shRect.Top
            ' Set the area of the clipping region (in pixels).
        If XHi > 310 Then XHi = 320
        If YHi > 310 Then YHi = 320
            picClip.ClipWidth = (Abs(XHi \ 10 - XLo \ 10))  'shRect.Width
            PicClipMove.ClipWidth = shRect.Width
            picClip.ClipHeight = (Abs(YHi \ 10 - YLo \ 10)) 'shRect.Height
            PicClipMove.ClipHeight = shRect.Height
            If shRect.Width < 2 Or shRect.Height < 2 Then
               MsgBox "Select Area using Click and Drag. After Area is selected, right click in selection for clipboard functions or use Edit Menu."
                cmdRegion_Click
                canSelect = False
                Exit Sub
            End If
            picMove.Width = shRect.Width
            picMove.Height = shRect.Height
            picMove.PaintPicture PicClipMove.Clip, 0, 0
            picMove.Left = shRect.Left
            picMove.Top = shRect.Top
            picMove.Visible = True
            xDelLo = shRect.Left \ 10
            xdelHi = (shRect.Left + shRect.Width) \ 10 - 1
            yDelLo = shRect.Top \ 10
            ydelHi = (shRect.Top + shRect.Height) \ 10 - 1
            '=======
            xOff = shRect.Left
            yOff = shRect.Top
            selectIt = False
    End If
            canSelect = False
            shRect.Move shRect.Left - 1, shRect.Top - 1, shRect.Width + 2, shRect.Height + 2
            Exit Sub
End If
'=========Draw a Line=======================
If lineDraw Then
    Dirty = True
    picContainer.DrawWidth = 10
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
    If lineOKDraw Then
        'picContainer.Line (lineX1, lineY1)-(X, Y), cr
        picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr
        'color last pixel
        picReal.PSet (X \ 10, Y \ 10), cr
        
        UpdateUndo
        
    End If
        picContainer.DrawWidth = 1
        DoEvents
        picReal.Picture = picReal.Image
        PaintDown
        Form1.Refresh
        lineOKDraw = False
        Line2.Visible = False
        Exit Sub
End If
'=========Draw a Rectangle or Fill Box=======================
If rectDraw Or fillBoxDraw Then
                Dirty = True
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
    If rectDraw Then
        picContainer.DrawWidth = 10
        Shape1.FillStyle = 1
        DoEvents
           ' picContainer.Line (lineX1, lineY1)-(X, Y), cr, B
            picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr, B
         picContainer.DrawWidth = 1
         Shape1.Visible = False
         rectOKDraw = False
    End If
    If fillBoxDraw Then
           ' picContainer.Line (lineX1, lineY1)-(X, Y), cr, BF
            picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr, BF
            Shape1.Visible = False
            rectOKDraw = False
    End If
       
       UpdateUndo
       
        DoEvents
        picReal.Picture = picReal.Image
        PaintDown
        Form1.Refresh
        Exit Sub
End If
'=========Draw a Circle or Filled Circle=======================
If circleDraw Or fillCircleDraw Then
                    Dirty = True
    If Button = 1 Then
            cr = PicMseColor.BackColor
    Else
            cr = PicMseColorR.BackColor
    End If
        picReal.FillColor = cr
        picContainer.FillColor = cr
    If circleDraw Then
        picContainer.DrawWidth = 10
        Shape1.FillStyle = 1
        DoEvents
            picReal.Circle (lineX1 \ 10, lineY1 \ 10), (Shape1.Width \ 10) / 2, cr, , , (Shape1.Height \ 10) / (Shape1.Width \ 10)
            picContainer.DrawWidth = 1
            Shape1.Visible = False
            circleOKDraw = False
    End If
    If fillCircleDraw Then
            picReal.Circle (lineX1 \ 10, lineY1 \ 10), (Shape1.Width \ 10) / 2, cr, , , (Shape1.Height \ 10) / (Shape1.Width \ 10)
            Shape1.Visible = False
            circleOKDraw = False
    End If
           UpdateUndo
           DoEvents
           picReal.Picture = picReal.Image
           PaintDown
           Form1.Refresh
           Exit Sub
End If
'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
    canDraw = False
    UpdateUndo
End If
Exit Sub
ExitIt: If circleOKDraw = True Then
Shape1.Visible = False
circleOKDraw = False
End If
End Sub

Private Sub chkSave()
Dim reply
If Dirty = True Then
If LangA = "lgc" Then
reply = MsgBox("是否保存正在编辑的文件?", vbYesNo + vbInformation, "Save As")
Else
reply = MsgBox("Save Current Edited Icon?", vbYesNo + vbInformation, "Save As")
End If
    
        If reply = vbYes Then mnuSave_Click
End If
End Sub

Private Sub PicMseColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMore_MouseDown 1, 0, 1, 1
End Sub

Private Sub PicMseColorR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMore_MouseDown 2, 0, 1, 1
End Sub

Private Sub picReal_Click()
picReal16.Picture = LoadPicture()
picReal.Picture = picReal.Image
picReal16.PaintPicture picReal.Picture, 0, 0, 16, 16
End Sub

Private Sub picReal16_Click()
Label6_Click
End Sub

Private Sub picTransparent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
PicMseColor.BackColor = picTransparent.Point(X, Y)
Else
PicMseColorR.BackColor = picTransparent.Point(X, Y)
End If
End Sub

Private Sub picTransparent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = picTransparent.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr
End Sub

Private Sub picMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPopUp
Else
    If LangA = "lgc" Then
MsgBox "弹出菜单请单击右键", vbInformation
Else
MsgBox "Right click for PopUp menu or use Edit on Menubar.", vbInformation
End If
    
End If
End Sub

Private Sub PicMseColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = PicMseColor.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr
End Sub

Private Sub PicMseColorR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = PicMseColorR.Point(X, Y)
r = clr Mod 256
G = (clr \ 256) Mod 256
B = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & G & " B " & B & "   -   " & clr
End Sub

Private Sub ExtractRequest()
Dim Answ, cdDir, cdIndex, Pos
   Dim hImgLarge As Long
   Dim hImgSmall As Long   'the handle to the system image list
   Dim fName As String     'the file name to get icon from
   Dim r As Long
fName$ = filePath 'Command$
   
   On Local Error GoTo cmdLoadErrorHandler
   
   hImgSmall& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

   hImgLarge& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   
   picTest.Picture = LoadPicture()
   picTest.AutoRedraw = True

   picTest.BackColor = RGB(197, 197, 197)
  'draw the associated icon into the pictureboxes
   Call ImageList_Draw(hImgLarge&, shinfo.iIcon, picReal.hdc, 0, 0, ILD_TRANSPARENT)
PaintDown
Form1.Refresh
MousePointer = 0
UpdateUndo
cmdUndo.Visible = False
Toolbar1.Buttons(5).Enabled = False
lblPath.Caption = filePath 'Command$
Exit Sub

cmdLoadErrorHandler:
MsgBox "Error # " & Err.Number & " - " & Err.Description, vbInformation
Exit Sub
End Sub

Private Sub setSwitchesFalse()
okToMove = False
eraseIt = False
rectDraw = False
lineDraw = False
floodDraw = False
fillBoxDraw = False
circleDraw = False
fillCircleDraw = False
pixelDraw = False
textDraw = False
selRegion = False
Timer1.Enabled = False
shRect.Visible = False
canSelect = False
moveIt = False
selectIt = False
picMove.Visible = False
picMove.Picture = LoadPicture()
picReal16.Picture = LoadPicture()
picContainer.Cls
PaintDown
picContainer.Picture = picContainer.Image
Form1.Refresh
End Sub


Private Sub Timer1_Timer()
If shRect.Borderstyle = vbBSDot Then
    shRect.Borderstyle = vbBSDashDot
            Else
                shRect.Borderstyle = vbBSDot
            End If
End Sub
' Return the long file name for a short file name.
Public Function LongFileName(ByVal short_name As String) As String
Dim Pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        Pos = 3
    Else
        result = ""
        Pos = 1
    End If

    ' Consider each section in the file name.
    Do While Pos > 0
        ' Find the next \.
        Pos = InStr(Pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If Pos = 0 Then
            long_name = Dir$(short_name, vbNormal + vbHidden + vbSystem + vbDirectory)
        Else
            long_name = Dir$(Left$(short_name, Pos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
Case "keyNew"
mnuNew_Click
picContainer.SetFocus

Case "keyOpen"
mnuOpen_Click
picContainer.SetFocus

Case "keySave"
mnuSave_Click
picContainer.SetFocus

Case "keyUndo"
cmdUndo_Click

Case "keyRedo"
cmdRedo_Click

Case "keyCut"
MnuEditOpts_Click (3)

Case "keyCopy"
MnuEditOpts_Click (4)

Case "keyPaste"
MnuEditOpts_Click (5)

Case "keyDel"
MnuEditOpts_Click (7)

Case "keyCS"
mnuPopCancel_Click

Case "keyPlugin"
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "Cannot Find Plugins Application", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "cs", App.Path, 1
End If

End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "keyOpen1"
mnuOpen_Click

Case "keyOpen2"
mnuExtract_Click

Case "keyCS2"
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "Cannot Find Plugins Application", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "cs", App.Path, 1
End If
End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReEnToolbar
End Sub
