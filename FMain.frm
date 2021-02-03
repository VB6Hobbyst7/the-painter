VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   Caption         =   "AP"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   -1650
   ClientWidth     =   11880
   ForeColor       =   &H00C00000&
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   7560
      Top             =   1200
   End
   Begin VB.Frame FrDS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   10440
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   3840
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8175
         Index           =   6
         Left            =   4440
         TabIndex        =   247
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6975
            Index           =   11
            Left            =   140
            TabIndex        =   248
            Top             =   480
            Width           =   3560
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   79
               Left            =   240
               Picture         =   "FMain.frx":0CCA
               Stretch         =   -1  'True
               Top             =   2610
               Width           =   690
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   79
               Left            =   960
               TabIndex        =   263
               Top             =   2760
               Width           =   2415
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   78
               Left            =   2040
               Picture         =   "FMain.frx":7558
               Stretch         =   -1  'True
               Top             =   5160
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   78
               Left            =   1800
               TabIndex        =   261
               Top             =   6360
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   77
               Left            =   240
               Picture         =   "FMain.frx":DDE6
               Stretch         =   -1  'True
               Top             =   5160
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   77
               Left            =   0
               TabIndex        =   259
               Top             =   6360
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   76
               Left            =   2040
               Picture         =   "FMain.frx":14674
               Stretch         =   -1  'True
               Top             =   3600
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   76
               Left            =   1800
               TabIndex        =   257
               Top             =   4800
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   75
               Left            =   240
               Picture         =   "FMain.frx":1AF02
               Stretch         =   -1  'True
               Top             =   3600
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   75
               Left            =   0
               TabIndex        =   255
               Top             =   4800
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   74
               Left            =   240
               Picture         =   "FMain.frx":21790
               Stretch         =   -1  'True
               Top             =   1770
               Width           =   690
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   74
               Left            =   960
               TabIndex        =   253
               Top             =   1920
               Width           =   2415
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   73
               Left            =   240
               Picture         =   "FMain.frx":2801E
               Stretch         =   -1  'True
               Top             =   930
               Width           =   690
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   73
               Left            =   960
               TabIndex        =   251
               Top             =   1080
               Width           =   2415
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   72
               Left            =   240
               Picture         =   "FMain.frx":2E8AC
               Stretch         =   -1  'True
               Top             =   90
               Width           =   690
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   72
               Left            =   960
               TabIndex        =   249
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  'Transparent
               Height          =   825
               Index           =   72
               Left            =   120
               TabIndex        =   250
               Top             =   0
               Width           =   3375
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  'Transparent
               Height          =   825
               Index           =   73
               Left            =   120
               TabIndex        =   252
               Top             =   840
               Width           =   3375
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  'Transparent
               Height          =   825
               Index           =   74
               Left            =   120
               TabIndex        =   254
               Top             =   1680
               Width           =   3375
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   75
               Left            =   0
               TabIndex        =   256
               Top             =   3600
               Width           =   1785
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   76
               Left            =   1800
               TabIndex        =   258
               Top             =   3600
               Width           =   1800
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   77
               Left            =   0
               TabIndex        =   260
               Top             =   5160
               Width           =   1785
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   78
               Left            =   1800
               TabIndex        =   262
               Top             =   5160
               Width           =   1800
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  'Transparent
               Height          =   825
               Index           =   79
               Left            =   120
               TabIndex        =   264
               Top             =   2520
               Width           =   3375
            End
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8350
         Index           =   5
         Left            =   4320
         TabIndex        =   219
         Top             =   1560
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   6735
            Index           =   9
            Left            =   120
            TabIndex        =   221
            Top             =   1560
            Width           =   3615
            Begin LP.Command Command17 
               Height          =   375
               Left            =   1440
               TabIndex        =   238
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":3513A
               Enabled         =   0   'False
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command18 
               Height          =   375
               Left            =   2520
               TabIndex        =   239
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":35156
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   70
               Left            =   1800
               TabIndex        =   236
               Top             =   5880
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   70
               Left            =   2040
               Picture         =   "FMain.frx":35172
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   69
               Left            =   0
               TabIndex        =   234
               Top             =   5880
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   69
               Left            =   240
               Picture         =   "FMain.frx":3BA00
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   68
               Left            =   1800
               TabIndex        =   232
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   68
               Left            =   2040
               Picture         =   "FMain.frx":4228E
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   67
               Left            =   0
               TabIndex        =   230
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   67
               Left            =   240
               Picture         =   "FMain.frx":48B1C
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   66
               Left            =   1800
               TabIndex        =   228
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   66
               Left            =   2040
               Picture         =   "FMain.frx":4F3AA
               Stretch         =   -1  'True
               Top             =   1560
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   65
               Left            =   0
               TabIndex        =   226
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   65
               Left            =   240
               Picture         =   "FMain.frx":55C38
               Stretch         =   -1  'True
               Top             =   1560
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   64
               Left            =   1800
               TabIndex        =   224
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   64
               Left            =   2040
               Picture         =   "FMain.frx":5C4C6
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   280
               Index           =   63
               Left            =   0
               TabIndex        =   222
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   63
               Left            =   240
               Picture         =   "FMain.frx":62D54
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   63
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   64
               Left            =   1800
               TabIndex        =   225
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   66
               Left            =   1800
               TabIndex        =   229
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   65
               Left            =   0
               TabIndex        =   227
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   67
               Left            =   0
               TabIndex        =   231
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   68
               Left            =   1800
               TabIndex        =   233
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   69
               Left            =   0
               TabIndex        =   235
               Top             =   4680
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   70
               Left            =   1800
               TabIndex        =   237
               Top             =   4680
               Width           =   1815
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6735
            Index           =   10
            Left            =   120
            TabIndex        =   240
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
            Begin LP.Command Command19 
               Height          =   375
               Left            =   1440
               TabIndex        =   243
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":695E2
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command20 
               Height          =   375
               Left            =   2520
               TabIndex        =   244
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":695FE
               Enabled         =   0   'False
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   280
               Index           =   71
               Left            =   0
               TabIndex        =   241
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   71
               Left            =   240
               Picture         =   "FMain.frx":6961A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   71
               Left            =   0
               TabIndex        =   242
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.Image ImgTi 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   3
            Left            =   480
            Picture         =   "FMain.frx":6FEA8
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label LabTi 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   4
            Left            =   1800
            TabIndex        =   220
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8350
         Index           =   3
         Left            =   4200
         TabIndex        =   199
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6495
            Index           =   8
            Left            =   120
            TabIndex        =   201
            Top             =   1560
            Width           =   3615
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   61
               Left            =   1800
               TabIndex        =   216
               Top             =   5880
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   61
               Left            =   2040
               Picture         =   "FMain.frx":76736
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   60
               Left            =   0
               TabIndex        =   214
               Top             =   5880
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   60
               Left            =   240
               Picture         =   "FMain.frx":7CFC4
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   59
               Left            =   1800
               TabIndex        =   212
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   59
               Left            =   2040
               Picture         =   "FMain.frx":83852
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   58
               Left            =   0
               TabIndex        =   210
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   58
               Left            =   240
               Picture         =   "FMain.frx":8A0E0
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   57
               Left            =   1800
               TabIndex        =   208
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   57
               Left            =   2040
               Picture         =   "FMain.frx":9096E
               Stretch         =   -1  'True
               Top             =   1560
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   56
               Left            =   0
               TabIndex        =   206
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   56
               Left            =   240
               Picture         =   "FMain.frx":971FC
               Stretch         =   -1  'True
               Top             =   1560
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   62
               Left            =   2040
               Picture         =   "FMain.frx":9DA8A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   62
               Left            =   1800
               TabIndex        =   204
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   55
               Left            =   240
               Picture         =   "FMain.frx":A4318
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   280
               Index           =   55
               Left            =   0
               TabIndex        =   202
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   55
               Left            =   0
               TabIndex        =   203
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   62
               Left            =   1800
               TabIndex        =   205
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   56
               Left            =   0
               TabIndex        =   207
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   57
               Left            =   1800
               TabIndex        =   209
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   58
               Left            =   0
               TabIndex        =   211
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   59
               Left            =   1800
               TabIndex        =   213
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   60
               Left            =   0
               TabIndex        =   215
               Top             =   4680
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   61
               Left            =   1800
               TabIndex        =   217
               Top             =   4680
               Width           =   1815
            End
         End
         Begin VB.Label LabTi 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   3
            Left            =   1800
            TabIndex        =   200
            Top             =   600
            Width           =   1695
         End
         Begin VB.Image ImgTi 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   2
            Left            =   480
            Picture         =   "FMain.frx":AABA6
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8350
         Index           =   4
         Left            =   4080
         TabIndex        =   184
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C185C5&
            Height          =   6135
            Index           =   7
            Left            =   120
            TabIndex        =   186
            Top             =   1920
            Width           =   3615
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   54
               Left            =   2040
               Picture         =   "FMain.frx":B1436
               Stretch         =   -1  'True
               Top             =   4080
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   465
               Index           =   54
               Left            =   1800
               TabIndex        =   197
               Top             =   5280
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   53
               Left            =   240
               Picture         =   "FMain.frx":B7CC4
               Stretch         =   -1  'True
               Top             =   4080
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   53
               Left            =   0
               TabIndex        =   195
               Top             =   5280
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   52
               Left            =   2040
               Picture         =   "FMain.frx":BE552
               Stretch         =   -1  'True
               Top             =   2040
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   465
               Index           =   52
               Left            =   1800
               TabIndex        =   193
               Top             =   3240
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   51
               Left            =   240
               Picture         =   "FMain.frx":C4DE0
               Stretch         =   -1  'True
               Top             =   2040
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   51
               Left            =   0
               TabIndex        =   191
               Top             =   3240
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   465
               Index           =   50
               Left            =   1800
               TabIndex        =   188
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   50
               Left            =   2040
               Picture         =   "FMain.frx":CB66E
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   49
               Left            =   0
               TabIndex        =   187
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   49
               Left            =   240
               Picture         =   "FMain.frx":D1EFC
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   49
               Left            =   0
               TabIndex        =   189
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   50
               Left            =   1800
               TabIndex        =   190
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   51
               Left            =   0
               TabIndex        =   192
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   52
               Left            =   1800
               TabIndex        =   194
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   53
               Left            =   0
               TabIndex        =   196
               Top             =   4080
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1665
               Index           =   54
               Left            =   1800
               TabIndex        =   198
               Top             =   4080
               Width           =   1815
            End
         End
         Begin VB.Image ImgTi 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   1
            Left            =   480
            Picture         =   "FMain.frx":D878A
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label LabTi 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   2
            Left            =   1800
            TabIndex        =   185
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8350
         Index           =   2
         Left            =   3960
         TabIndex        =   134
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6735
            Index           =   6
            Left            =   120
            TabIndex        =   170
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
            Begin LP.Command Command15 
               Height          =   375
               Left            =   1440
               TabIndex        =   181
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":DF01A
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command16 
               Height          =   375
               Left            =   2520
               TabIndex        =   182
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":DF036
               Enabled         =   0   'False
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   45
               Left            =   2040
               Picture         =   "FMain.frx":DF052
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   45
               Left            =   1800
               TabIndex        =   177
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   44
               Left            =   240
               Picture         =   "FMain.frx":E58E0
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   44
               Left            =   0
               TabIndex        =   175
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   47
               Left            =   2040
               Picture         =   "FMain.frx":EC16E
               Stretch         =   -1  'True
               Top             =   1815
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   47
               Left            =   1800
               TabIndex        =   173
               Top             =   3015
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   46
               Left            =   240
               Picture         =   "FMain.frx":F29FC
               Stretch         =   -1  'True
               Top             =   1815
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Index           =   46
               Left            =   0
               TabIndex        =   171
               Top             =   3015
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1740
               Index           =   47
               Left            =   1800
               TabIndex        =   174
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1740
               Index           =   46
               Left            =   0
               TabIndex        =   172
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1740
               Index           =   44
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1740
               Index           =   45
               Left            =   1800
               TabIndex        =   178
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6735
            Index           =   5
            Left            =   120
            TabIndex        =   155
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
            Begin LP.Command Command13 
               Height          =   375
               Left            =   1440
               TabIndex        =   168
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":F928A
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command14 
               Height          =   375
               Left            =   2520
               TabIndex        =   169
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":F92A6
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1215
               Index           =   48
               Left            =   480
               Picture         =   "FMain.frx":F92C2
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   2685
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   48
               Left            =   360
               TabIndex        =   179
               Top             =   5895
               Width           =   2895
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   43
               Left            =   2040
               Picture         =   "FMain.frx":106500
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   43
               Left            =   1800
               TabIndex        =   166
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   42
               Left            =   240
               Picture         =   "FMain.frx":10CD8E
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   42
               Left            =   0
               TabIndex        =   164
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   41
               Left            =   2040
               Picture         =   "FMain.frx":11361C
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   41
               Left            =   1800
               TabIndex        =   162
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   40
               Left            =   240
               Picture         =   "FMain.frx":119EAA
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   40
               Left            =   0
               TabIndex        =   160
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   39
               Left            =   2040
               Picture         =   "FMain.frx":120738
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   39
               Left            =   1800
               TabIndex        =   158
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   38
               Left            =   240
               Picture         =   "FMain.frx":126FC6
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   38
               Left            =   0
               TabIndex        =   156
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   38
               Left            =   0
               TabIndex        =   157
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   39
               Left            =   1800
               TabIndex        =   159
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   40
               Left            =   0
               TabIndex        =   161
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   41
               Left            =   1800
               TabIndex        =   163
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   42
               Left            =   0
               TabIndex        =   165
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   43
               Left            =   1800
               TabIndex        =   167
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   48
               Left            =   0
               TabIndex        =   180
               Top             =   4680
               Width           =   3615
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6735
            Index           =   4
            Left            =   120
            TabIndex        =   136
            Top             =   1560
            Width           =   3615
            Begin LP.Command Command11 
               Height          =   375
               Left            =   1440
               TabIndex        =   153
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":12D854
               Enabled         =   0   'False
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command12 
               Height          =   375
               Left            =   2520
               TabIndex        =   154
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":12D870
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1255
               Index           =   37
               Left            =   2040
               Picture         =   "FMain.frx":12D88C
               Stretch         =   -1  'True
               Top             =   4655
               Width           =   1410
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   37
               Left            =   1800
               TabIndex        =   151
               Top             =   5925
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1230
               Index           =   36
               Left            =   240
               Picture         =   "FMain.frx":13411A
               Stretch         =   -1  'True
               Top             =   4659
               Width           =   1380
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   36
               Left            =   0
               TabIndex        =   149
               Top             =   5920
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1150
               Index           =   35
               Left            =   2040
               Picture         =   "FMain.frx":13A9A8
               Stretch         =   -1  'True
               Top             =   3125
               Width           =   1350
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   35
               Left            =   1800
               TabIndex        =   147
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1155
               Index           =   34
               Left            =   240
               Picture         =   "FMain.frx":141236
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   34
               Left            =   0
               TabIndex        =   145
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   33
               Left            =   2040
               Picture         =   "FMain.frx":147AC4
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   33
               Left            =   1800
               TabIndex        =   143
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   32
               Left            =   240
               Picture         =   "FMain.frx":14E352
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   32
               Left            =   0
               TabIndex        =   141
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   31
               Left            =   2040
               Picture         =   "FMain.frx":154BE0
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   31
               Left            =   1800
               TabIndex        =   139
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   30
               Left            =   240
               Picture         =   "FMain.frx":15B46E
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
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
               Height          =   255
               Index           =   30
               Left            =   0
               TabIndex        =   137
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   30
               Left            =   0
               TabIndex        =   138
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   31
               Left            =   1800
               TabIndex        =   140
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   32
               Left            =   0
               TabIndex        =   142
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   33
               Left            =   1800
               TabIndex        =   144
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   34
               Left            =   0
               TabIndex        =   146
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   35
               Left            =   1800
               TabIndex        =   148
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1540
               Index           =   36
               Left            =   0
               TabIndex        =   150
               Top             =   4680
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   37
               Left            =   1800
               TabIndex        =   152
               Top             =   4680
               Width           =   1815
            End
         End
         Begin VB.Label LabTi 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   0
            Left            =   1800
            TabIndex        =   135
            Top             =   600
            Width           =   1695
         End
         Begin VB.Image ImgTi 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   0
            Left            =   480
            Picture         =   "FMain.frx":161CFC
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8350
         Index           =   1
         Left            =   3840
         TabIndex        =   96
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   6735
            Index           =   3
            Left            =   120
            TabIndex        =   117
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
            Begin LP.Command Command9 
               Height          =   375
               Left            =   1440
               TabIndex        =   132
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":16858C
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command10 
               Height          =   375
               Left            =   2520
               TabIndex        =   133
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":1685A8
               Enabled         =   0   'False
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   29
               Left            =   0
               TabIndex        =   130
               Top             =   5895
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   29
               Left            =   240
               Picture         =   "FMain.frx":1685C4
               Stretch         =   -1  'True
               Top             =   4695
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   28
               Left            =   1800
               TabIndex        =   128
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   28
               Left            =   2040
               Picture         =   "FMain.frx":16EE52
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   27
               Left            =   0
               TabIndex        =   126
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   27
               Left            =   240
               Picture         =   "FMain.frx":1756E0
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   26
               Left            =   1800
               TabIndex        =   124
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   26
               Left            =   2040
               Picture         =   "FMain.frx":17BF6E
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   25
               Left            =   0
               TabIndex        =   122
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   25
               Left            =   240
               Picture         =   "FMain.frx":1827FC
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   24
               Left            =   1800
               TabIndex        =   120
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   24
               Left            =   2040
               Picture         =   "FMain.frx":18908A
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   118
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   23
               Left            =   240
               Picture         =   "FMain.frx":18F918
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   24
               Left            =   1800
               TabIndex        =   121
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   23
               Left            =   0
               TabIndex        =   119
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   26
               Left            =   1800
               TabIndex        =   125
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   25
               Left            =   0
               TabIndex        =   123
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   28
               Left            =   1800
               TabIndex        =   129
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   27
               Left            =   0
               TabIndex        =   127
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   29
               Left            =   0
               TabIndex        =   131
               Top             =   4680
               Width           =   1815
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Height          =   6735
            Index           =   2
            Left            =   120
            TabIndex        =   98
            Top             =   1560
            Width           =   3615
            Begin LP.Command Command7 
               Height          =   375
               Left            =   1440
               TabIndex        =   115
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":1961A6
               Enabled         =   0   'False
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command8 
               Height          =   375
               Left            =   2520
               TabIndex        =   116
               Top             =   6240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":1961C2
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   22
               Left            =   1800
               TabIndex        =   113
               Top             =   5895
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   22
               Left            =   2040
               Picture         =   "FMain.frx":1961DE
               Stretch         =   -1  'True
               Top             =   4695
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   111
               Top             =   5895
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   21
               Left            =   240
               Picture         =   "FMain.frx":19CA6C
               Stretch         =   -1  'True
               Top             =   4695
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   20
               Left            =   1800
               TabIndex        =   109
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   20
               Left            =   2040
               Picture         =   "FMain.frx":1A32FA
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   107
               Top             =   4335
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   19
               Left            =   240
               Picture         =   "FMain.frx":1A9B88
               Stretch         =   -1  'True
               Top             =   3135
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   18
               Left            =   1800
               TabIndex        =   105
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   18
               Left            =   2040
               Picture         =   "FMain.frx":1B0416
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   103
               Top             =   2775
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   17
               Left            =   240
               Picture         =   "FMain.frx":1B6CA4
               Stretch         =   -1  'True
               Top             =   1575
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   16
               Left            =   1800
               TabIndex        =   101
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   16
               Left            =   2040
               Picture         =   "FMain.frx":1BD532
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   99
               Top             =   1215
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   15
               Left            =   240
               Picture         =   "FMain.frx":1C3DC0
               Stretch         =   -1  'True
               Top             =   15
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   15
               Left            =   0
               TabIndex        =   100
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   16
               Left            =   1800
               TabIndex        =   102
               Top             =   0
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   18
               Left            =   1800
               TabIndex        =   106
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   17
               Left            =   0
               TabIndex        =   104
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   20
               Left            =   1800
               TabIndex        =   110
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   19
               Left            =   0
               TabIndex        =   108
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   22
               Left            =   1800
               TabIndex        =   114
               Top             =   4680
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   21
               Left            =   0
               TabIndex        =   112
               Top             =   4680
               Width           =   1815
            End
         End
         Begin VB.Image ImgTi 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   15
            Left            =   480
            Picture         =   "FMain.frx":1CA64E
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label LabTi 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   1
            Left            =   1800
            TabIndex        =   97
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame ToolFunc 
         BackColor       =   &H00C185C5&
         Height          =   8175
         Index           =   0
         Left            =   3720
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame ToolF 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hard Colors"
            Height          =   5055
            Index           =   2
            Left            =   10680
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   3255
            Begin VB.Label CmdCloseTol 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   3000
               TabIndex        =   46
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hard Yellow"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   11
               Left            =   1440
               TabIndex        =   45
               Top             =   4200
               Width           =   1695
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hard Blue"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   10
               Left            =   1440
               TabIndex        =   44
               Top             =   3000
               Width           =   1695
            End
            Begin VB.Image ImgT2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   11
               Left            =   120
               Picture         =   "FMain.frx":1D0EDE
               Stretch         =   -1  'True
               Top             =   3840
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   10
               Left            =   120
               Picture         =   "FMain.frx":1D776C
               Stretch         =   -1  'True
               Top             =   2640
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   9
               Left            =   120
               Picture         =   "FMain.frx":1DDFFA
               Stretch         =   -1  'True
               Top             =   1440
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   8
               Left            =   120
               Picture         =   "FMain.frx":1E4888
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1290
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hard Green"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   9
               Left            =   1440
               TabIndex        =   43
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hard Red"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   8
               Left            =   1440
               TabIndex        =   42
               Top             =   600
               Width           =   1695
            End
         End
         Begin VB.Frame ToolF 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Soft Colors"
            Height          =   6255
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   1560
            Visible         =   0   'False
            Width           =   3375
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Purple"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   7
               Left            =   1440
               TabIndex        =   53
               Top             =   5400
               Width           =   1695
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Yellow"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   6
               Left            =   1440
               TabIndex        =   52
               Top             =   4200
               Width           =   1695
            End
            Begin VB.Image ImgT2 
               Height          =   1125
               Index           =   7
               Left            =   120
               Picture         =   "FMain.frx":1EB116
               Stretch         =   -1  'True
               Top             =   5040
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Height          =   1125
               Index           =   6
               Left            =   120
               Picture         =   "FMain.frx":1F19A4
               Stretch         =   -1  'True
               Top             =   3840
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Height          =   1125
               Index           =   5
               Left            =   120
               Picture         =   "FMain.frx":1F8232
               Stretch         =   -1  'True
               Top             =   2640
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Height          =   1125
               Index           =   4
               Left            =   120
               Picture         =   "FMain.frx":1FEAC0
               Stretch         =   -1  'True
               Top             =   1440
               Width           =   1290
            End
            Begin VB.Image ImgT2 
               Height          =   1125
               Index           =   3
               Left            =   120
               Picture         =   "FMain.frx":20534E
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1290
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Orange"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   5
               Left            =   1440
               TabIndex        =   51
               Top             =   3000
               Width           =   1695
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Green"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   4
               Left            =   1440
               TabIndex        =   50
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Red"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   3
               Left            =   1440
               TabIndex        =   49
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label CmdCloseTol 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   3000
               TabIndex        =   48
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame ToolF 
            BackColor       =   &H00FFC0C0&
            Caption         =   "特殊浮雕"
            Height          =   4575
            Index           =   0
            Left            =   240
            TabIndex        =   54
            Top             =   1800
            Visible         =   0   'False
            Width           =   3375
            Begin VB.Label CmdCloseTol 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   58
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hold Red"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   2
               Left            =   1680
               TabIndex        =   57
               Top             =   3600
               Width           =   1455
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hold Green"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   735
               Index           =   1
               Left            =   1680
               TabIndex        =   56
               Top             =   2160
               Width           =   1455
            End
            Begin VB.Label Leffect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Hold Red"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   495
               Index           =   0
               Left            =   1680
               TabIndex        =   55
               Top             =   720
               Width           =   1455
            End
            Begin VB.Image ImgT2 
               Height          =   1335
               Index           =   2
               Left            =   120
               Picture         =   "FMain.frx":20BBDC
               Top             =   3120
               Width           =   1500
            End
            Begin VB.Image ImgT2 
               Height          =   1335
               Index           =   1
               Left            =   120
               Picture         =   "FMain.frx":21246A
               Top             =   1680
               Width           =   1500
            End
            Begin VB.Image ImgT2 
               Height          =   1335
               Index           =   0
               Left            =   120
               Picture         =   "FMain.frx":218CF8
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   6465
            Index           =   1
            Left            =   4080
            TabIndex        =   76
            Top             =   1440
            Width           =   3615
            Begin LP.Command Command5 
               Height          =   375
               Left            =   1440
               TabIndex        =   77
               Top             =   6000
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":21F586
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command6 
               Height          =   375
               Left            =   2520
               TabIndex        =   78
               Top             =   6000
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":21F5A2
               Enabled         =   0   'False
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   12
               Left            =   2160
               Picture         =   "FMain.frx":21F5BE
               Stretch         =   -1  'True
               Top             =   1320
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   11
               Left            =   240
               Picture         =   "FMain.frx":225E4C
               Stretch         =   -1  'True
               Top             =   2880
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   10
               Left            =   2160
               Picture         =   "FMain.frx":22C6DA
               Stretch         =   -1  'True
               Top             =   2880
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   9
               Left            =   240
               Picture         =   "FMain.frx":232F68
               Stretch         =   -1  'True
               Top             =   4440
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   85
               Top             =   4080
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   10
               Left            =   1920
               TabIndex        =   84
               Top             =   4080
               Width           =   1695
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   83
               Top             =   5640
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   8
               Left            =   1920
               TabIndex        =   82
               Top             =   5640
               Width           =   1695
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   13
               Left            =   240
               Picture         =   "FMain.frx":2397F6
               Stretch         =   -1  'True
               Top             =   1320
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   12
               Left            =   1920
               TabIndex        =   81
               Top             =   2520
               Width           =   1695
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   8
               Left            =   2160
               Picture         =   "FMain.frx":240084
               Stretch         =   -1  'True
               Top             =   4440
               Width           =   1290
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   80
               Top             =   2520
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Soft Colors..."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   375
               Index           =   14
               Left            =   1800
               TabIndex        =   79
               Top             =   405
               Width           =   1695
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   14
               Left            =   480
               Picture         =   "FMain.frx":246912
               Stretch         =   -1  'True
               Top             =   45
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1215
               Index           =   14
               Left            =   0
               TabIndex        =   92
               Top             =   0
               Width           =   3615
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1580
               Index           =   13
               Left            =   0
               TabIndex        =   91
               Top             =   1245
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1575
               Index           =   12
               Left            =   1920
               TabIndex        =   90
               Top             =   1250
               Width           =   1695
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   8
               Left            =   1920
               TabIndex        =   89
               Top             =   4400
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1550
               Index           =   9
               Left            =   0
               TabIndex        =   88
               Top             =   4400
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   11
               Left            =   0
               TabIndex        =   87
               Top             =   2860
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   10
               Left            =   1920
               TabIndex        =   86
               Top             =   2860
               Width           =   1695
            End
         End
         Begin VB.Frame Tol1 
            BackColor       =   &H00C185C5&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   6465
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   3615
            Begin LP.Command Command3 
               Height          =   375
               Left            =   1440
               TabIndex        =   60
               Top             =   6000
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":24D1A0
               Enabled         =   0   'False
               Caption         =   "<< Back"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin LP.Command Command4 
               Height          =   375
               Left            =   2520
               TabIndex        =   61
               Top             =   6000
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Icon            =   "FMain.frx":24D1BC
               Caption         =   "Next >>"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "SimSun"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   7
               Left            =   1920
               TabIndex        =   68
               Top             =   5640
               Width           =   1695
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   67
               Top             =   5640
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   5
               Left            =   1920
               TabIndex        =   66
               Top             =   4080
               Width           =   1695
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   65
               Top             =   4080
               Width           =   1815
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss Special"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   615
               Index           =   3
               Left            =   1800
               TabIndex        =   64
               Top             =   2040
               Width           =   1575
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Engrave"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   2
               Left            =   1920
               TabIndex        =   63
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label LabT1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Emboss"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   62
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   3
               Left            =   480
               Picture         =   "FMain.frx":24D1D8
               Stretch         =   -1  'True
               Top             =   1680
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   7
               Left            =   2160
               Picture         =   "FMain.frx":253A66
               Stretch         =   -1  'True
               Top             =   4440
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   6
               Left            =   240
               Picture         =   "FMain.frx":25A2F4
               Stretch         =   -1  'True
               Top             =   4440
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   5
               Left            =   2160
               Picture         =   "FMain.frx":260B82
               Stretch         =   -1  'True
               Top             =   2880
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   4
               Left            =   120
               Picture         =   "FMain.frx":267410
               Stretch         =   -1  'True
               Top             =   2880
               Width           =   1530
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   2
               Left            =   2160
               Picture         =   "FMain.frx":26DC9E
               Stretch         =   -1  'True
               Top             =   120
               Width           =   1290
            End
            Begin VB.Image ImgT1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1125
               Index           =   1
               Left            =   240
               Picture         =   "FMain.frx":27452C
               Stretch         =   -1  'True
               Top             =   120
               Width           =   1290
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1575
               Index           =   1
               Left            =   0
               TabIndex        =   75
               Top             =   45
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1575
               Index           =   2
               Left            =   1920
               TabIndex        =   74
               Top             =   50
               Width           =   1695
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   4
               Left            =   0
               TabIndex        =   73
               Top             =   2860
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1500
               Index           =   5
               Left            =   1920
               TabIndex        =   72
               Top             =   2860
               Width           =   1695
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1550
               Index           =   6
               Left            =   0
               TabIndex        =   71
               Top             =   4400
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1545
               Index           =   7
               Left            =   1920
               TabIndex        =   70
               Top             =   4400
               Width           =   1815
            End
            Begin VB.Label SelectTool 
               BackColor       =   &H00DCB8DA&
               BackStyle       =   0  'Transparent
               Height          =   1215
               Index           =   3
               Left            =   0
               TabIndex        =   69
               Top             =   1630
               Width           =   3615
            End
         End
         Begin VB.Label LabT1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   495
            Index           =   0
            Left            =   1800
            TabIndex        =   94
            Top             =   600
            Width           =   1695
         End
         Begin VB.Image ImgT1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Index           =   0
            Left            =   360
            Picture         =   "FMain.frx":27ADBA
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label SelectTool 
            BackStyle       =   0  'Transparent
            Caption         =   "LP"
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   93
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin LP.Command CmdBack 
         Height          =   495
         Left            =   2040
         TabIndex        =   95
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Icon            =   "FMain.frx":28164A
         Caption         =   "<< 返回列表"
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
      Begin LP.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   1080
         TabIndex        =   32
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":281666
         Caption         =   "一般滤镜"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":281682
         Caption         =   "特殊滤镜"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   1080
         TabIndex        =   34
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":28169E
         Caption         =   "修饰"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   1080
         TabIndex        =   35
         Top             =   5760
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":2816BA
         Caption         =   "变形"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   1080
         TabIndex        =   36
         Top             =   4440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":2816D6
         Caption         =   "图片"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   5
         Left            =   1080
         TabIndex        =   183
         Top             =   7080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":2816F2
         Caption         =   "颜色"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin LP.Command Command2 
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   218
         Top             =   8400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Icon            =   "FMain.frx":28170E
         Caption         =   "填充"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontColor       =   16711680
         ButtonLook      =   4
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   6
         Left            =   480
         Picture         =   "FMain.frx":28172A
         Top             =   8400
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   5
         Left            =   480
         Picture         =   "FMain.frx":2823F4
         Top             =   7080
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   4
         Left            =   480
         Picture         =   "FMain.frx":2830BE
         Top             =   4440
         Width           =   480
      End
      Begin VB.Label ToolName 
         BackStyle       =   0  'Transparent
         Caption         =   "ToolName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   780
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label ToolName2 
         BackStyle       =   0  'Transparent
         Caption         =   "ToolName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1350
         TabIndex        =   38
         Top             =   810
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Image ToolUseImg 
         Height          =   480
         Left            =   600
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   3
         Left            =   480
         Picture         =   "FMain.frx":283D88
         Top             =   5760
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   2
         Left            =   480
         Picture         =   "FMain.frx":284A52
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "FMain.frx":28571C
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Index           =   0
         Left            =   480
         Picture         =   "FMain.frx":2863E6
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   22440
         Left            =   0
         Picture         =   "FMain.frx":2870B0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8085
      End
   End
   Begin VB.CommandButton FCS 
      Caption         =   "Command3"
      Height          =   255
      Left            =   6240
      TabIndex        =   31
      Top             =   1.66920e5
      Width           =   1095
   End
   Begin LP.Command Command1 
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   5760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Icon            =   "FMain.frx":290A05
      Caption         =   "&Preferences..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   6000
      Left            =   6000
      Top             =   480
   End
   Begin VB.CommandButton Run2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "打开一个图片"
      DownPicture     =   "FMain.frx":290A21
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4680
      MouseIcon       =   "FMain.frx":292563
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2926B5
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1920
      Top             =   2760
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   7320
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Tempmem2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   765
      MousePointer    =   2  'Cross
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   11
      Top             =   1995
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1395
      Top             =   2760
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   315
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":2941AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":294313
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   5040
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "还原操作"
            Key             =   "keyUndo"
            Object.ToolTipText     =   "还原"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "撤消选择"
            Key             =   "keySelectAll"
            Object.ToolTipText     =   "撤消选择"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox TempMem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   315
      MousePointer    =   2  'Cross
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   8
      Top             =   1995
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   900
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox PicX 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   9120
      ScaleHeight     =   553
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   550
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   8250
      Begin VB.PictureBox Dummy 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7080
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   5
         Top             =   7080
         Width           =   225
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         LargeChange     =   10
         Left            =   0
         TabIndex        =   4
         Top             =   7065
         Width           =   7080
      End
      Begin VB.VScrollBar VS1 
         Height          =   7080
         LargeChange     =   10
         Left            =   7080
         TabIndex        =   3
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Text 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   405
         TabIndex        =   13
         Top             =   1890
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Text 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   405
         TabIndex        =   12
         Top             =   1665
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   915
         Left            =   360
         Top             =   405
         Width           =   825
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   7920
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1032
            MinWidth        =   1032
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "Ins"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:10 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Run1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run1"
      DownPicture     =   "FMain.frx":294473
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4680
      MouseIcon       =   "FMain.frx":2A04B5
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2A0607
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "点这里感受更加直观、快速、方便的图片编辑操作"
      Height          =   375
      Left            =   5400
      MouseIcon       =   "FMain.frx":2A2221
      MousePointer    =   99  'Custom
      TabIndex        =   246
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "更加直观地编辑图片"
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
      Left            =   5640
      MouseIcon       =   "FMain.frx":2A2373
      MousePointer    =   99  'Custom
      TabIndex        =   245
      Top             =   1350
      Width           =   2535
   End
   Begin VB.Line Line4 
      X1              =   288
      X2              =   336
      Y1              =   168
      Y2              =   144
   End
   Begin VB.Line Line3 
      X1              =   288
      X2              =   328
      Y1              =   168
      Y2              =   128
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   5160
      Picture         =   "FMain.frx":2A24C5
      Top             =   1320
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Image IBd3 
      Height          =   1905
      Left            =   4200
      Picture         =   "FMain.frx":2A2807
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image IBd2 
      Height          =   1905
      Left            =   4200
      MouseIcon       =   "FMain.frx":2A4411
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2A4563
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image IBd1 
      Height          =   1905
      Left            =   4200
      MouseIcon       =   "FMain.frx":2A616D
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2A62BF
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image IBg3 
      Height          =   1905
      Left            =   4200
      Picture         =   "FMain.frx":2A7EC9
      Top             =   360
      Width           =   270
   End
   Begin VB.Image IBg2 
      Height          =   1905
      Left            =   4200
      MouseIcon       =   "FMain.frx":2A9AD3
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2A9C25
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image IBg1 
      Height          =   1905
      Left            =   4200
      MouseIcon       =   "FMain.frx":2AB82F
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":2AB981
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Studio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   90
      Width           =   1095
   End
   Begin VB.Image Bover 
      Height          =   300
      Left            =   120
      Picture         =   "FMain.frx":2AD58B
      Top             =   45
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Bmove 
      Height          =   300
      Left            =   120
      Picture         =   "FMain.frx":2AED3D
      Top             =   45
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image PicNote1 
      Height          =   240
      Left            =   2400
      Picture         =   "FMain.frx":2B04EF
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LabNote1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "FMain.frx":2B0EF1
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Image Image22 
      Height          =   225
      Left            =   1965
      Picture         =   "FMain.frx":2B32C3
      Top             =   4200
      Width           =   225
   End
   Begin VB.Image Image21 
      Height          =   225
      Left            =   480
      Picture         =   "FMain.frx":2B3633
      Top             =   4200
      Width           =   225
   End
   Begin VB.Image HelpImg1 
      Height          =   360
      Left            =   3600
      Picture         =   "FMain.frx":2B3996
      ToolTipText     =   "怎么看坐标？"
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(304)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   5
      Left            =   960
      MouseIcon       =   "FMain.frx":2B4098
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(303)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   4
      Left            =   960
      MouseIcon       =   "FMain.frx":2B41EA
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(302)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   3
      Left            =   960
      MouseIcon       =   "FMain.frx":2B433C
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(301)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   2
      Left            =   960
      MouseIcon       =   "FMain.frx":2B448E
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(300)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   1
      Left            =   960
      MouseIcon       =   "FMain.frx":2B45E0
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(299)"
      Enabled         =   0   'False
      ForeColor       =   &H00404080&
      Height          =   255
      Index           =   0
      Left            =   960
      MouseIcon       =   "FMain.frx":2B4732
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   5
      Left            =   480
      Picture         =   "FMain.frx":2B4884
      Top             =   7050
      Width           =   240
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   4
      Left            =   480
      Picture         =   "FMain.frx":2B4B4D
      Top             =   6690
      Width           =   240
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   3
      Left            =   480
      Picture         =   "FMain.frx":2B4E16
      Top             =   5970
      Width           =   240
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "FMain.frx":2B50E2
      Top             =   6330
      Width           =   240
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "FMain.frx":2B53AB
      Top             =   5610
      Width           =   240
   End
   Begin VB.Image Image20 
      Height          =   240
      Index           =   0
      Left            =   480
      Picture         =   "FMain.frx":2B5677
      Top             =   5250
      Width           =   240
   End
   Begin VB.Image Image19 
      Height          =   45
      Left            =   120
      Picture         =   "FMain.frx":2B5943
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   4065
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(298)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   20
      Top             =   4830
      Width           =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(297)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   19
      Top             =   3510
      Width           =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(296)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   17
      Top             =   1710
      Width           =   1680
   End
   Begin VB.Image Image16 
      Height          =   45
      Left            =   120
      Picture         =   "FMain.frx":2B62F1
      Stretch         =   -1  'True
      Top             =   3330
      Width           =   4065
   End
   Begin VB.Image Image13 
      Height          =   45
      Left            =   120
      Picture         =   "FMain.frx":2B6C9F
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4065
   End
   Begin VB.Image Image11 
      Height          =   45
      Left            =   120
      Picture         =   "FMain.frx":2B764D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   4080
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   120
      Picture         =   "FMain.frx":2B7FFB
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   4065
   End
   Begin VB.Image Image6 
      Height          =   330
      Left            =   120
      Picture         =   "FMain.frx":2BC765
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   4065
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   120
      Picture         =   "FMain.frx":2C0ECF
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4065
   End
   Begin VB.Image SideBg 
      Height          =   15360
      Left            =   4185
      Picture         =   "FMain.frx":2C5639
      Stretch         =   -1  'True
      Top             =   360
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "信息"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lgT(295)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   390
      Width           =   1785
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "坐标"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   510
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   3480
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   4
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   735
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   2565
      Stretch         =   -1  'True
      Top             =   735
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   1770
      Stretch         =   -1  'True
      Top             =   735
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   975
      Stretch         =   -1  'True
      Top             =   735
      Width           =   735
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   180
      Stretch         =   -1  'True
      Top             =   735
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "完成"
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
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   10200
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   120
      Picture         =   "FMain.frx":2D767B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4065
   End
   Begin VB.Image Image8 
      Enabled         =   0   'False
      Height          =   960
      Left            =   120
      Picture         =   "FMain.frx":2DBDE5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   45
   End
   Begin VB.Image Image9 
      Height          =   1440
      Left            =   120
      Picture         =   "FMain.frx":2DC607
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   45
   End
   Begin VB.Image Image15 
      Enabled         =   0   'False
      Height          =   960
      Left            =   120
      Picture         =   "FMain.frx":2DCE29
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   45
   End
   Begin VB.Image Image18 
      Height          =   2520
      Left            =   4140
      Picture         =   "FMain.frx":2DD64B
      Top             =   5040
      Width           =   45
   End
   Begin VB.Image Image17 
      Height          =   2520
      Left            =   120
      Picture         =   "FMain.frx":2DDE6D
      Top             =   5040
      Width           =   45
   End
   Begin VB.Image B1 
      Height          =   300
      Left            =   120
      Picture         =   "FMain.frx":2DE68F
      Top             =   45
      Width           =   1500
   End
   Begin VB.Image TabBg 
      Height          =   210
      Left            =   0
      Picture         =   "FMain.frx":2DF29F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   1280
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   1280
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   1280
      Y1              =   17
      Y2              =   17
   End
   Begin VB.Label Inject1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "软件简介："
      Height          =   1935
      Left            =   7080
      TabIndex        =   21
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Image Image14 
      Enabled         =   0   'False
      Height          =   960
      Left            =   4140
      Picture         =   "FMain.frx":2F3CE3
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   45
   End
   Begin VB.Image Image12 
      Height          =   1440
      Left            =   4140
      Picture         =   "FMain.frx":2F4505
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   45
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   960
      Left            =   4140
      Picture         =   "FMain.frx":2F4D27
      Stretch         =   -1  'True
      Top             =   600
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFiles 
         Caption         =   "打开文件..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "另存为文件..."
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "保存至""我的图片"""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "打印图片..."
         Index           =   4
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "启动向导..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Select Language..."
         Index           =   7
         Visible         =   0   'False
         Begin VB.Menu mnuSelLang 
            Caption         =   "En"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuSelLang 
            Caption         =   "SC"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "退出..."
         Index           =   8
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "编辑(&E)"
      Begin VB.Menu Cancel1 
         Caption         =   "撤消"
         Shortcut        =   ^Z
      End
      Begin VB.Menu h01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSel 
         Caption         =   "调整选择区域"
         Index           =   0
      End
      Begin VB.Menu mnuSel 
         Caption         =   "取消选择 (鼠标右键)"
         Index           =   1
      End
      Begin VB.Menu mnuSel 
         Caption         =   "全选"
         Index           =   2
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "图片(&P)"
      Begin VB.Menu mnuPic 
         Caption         =   "水平翻转图片"
         Index           =   0
      End
      Begin VB.Menu mnuPic 
         Caption         =   "垂直翻转图片"
         Index           =   1
      End
      Begin VB.Menu mnuPic 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPic 
         Caption         =   "左侧水平平分图象"
         Index           =   3
      End
      Begin VB.Menu mnuPic 
         Caption         =   "右侧水平平分图象"
         Index           =   4
      End
      Begin VB.Menu mnuPic 
         Caption         =   "上部垂直平分图象"
         Index           =   5
      End
      Begin VB.Menu mnuPic 
         Caption         =   "上部垂直平分图象"
         Index           =   6
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "颜色(&C)"
      Begin VB.Menu mnuCol 
         Caption         =   "调整颜色..."
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCol 
         Caption         =   "去除颜色"
         Index           =   2
         Begin VB.Menu mnuKill 
            Caption         =   "去除红色"
            Index           =   0
         End
         Begin VB.Menu mnuKill 
            Caption         =   "去除绿色"
            Index           =   1
         End
         Begin VB.Menu mnuKill 
            Caption         =   "去除蓝色"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCol 
         Caption         =   "转换基调"
         Index           =   4
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> BGR"
            Index           =   0
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> BRG"
            Index           =   1
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> GBR"
            Index           =   2
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> GRB"
            Index           =   3
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> RBG"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCol 
         Caption         =   "亮度..."
         Index           =   6
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuCol 
         Caption         =   "对比度..."
         Index           =   7
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCol 
         Caption         =   "反色"
         Index           =   9
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuCol 
         Caption         =   "转换颜色主调"
         Index           =   11
         Begin VB.Menu mnuInv 
            Caption         =   "转换红色调"
            Index           =   0
         End
         Begin VB.Menu mnuInv 
            Caption         =   "转换绿色调"
            Index           =   1
         End
         Begin VB.Menu mnuInv 
            Caption         =   "转换蓝色调"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCol 
         Caption         =   "灰度"
         Index           =   13
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "滤镜(&R)"
      Begin VB.Menu mnuFilter 
         Caption         =   "浮雕"
         Index           =   0
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "特殊浮雕"
         Index           =   1
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "突出红色"
            Index           =   0
         End
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "突出绿色"
            Index           =   1
         End
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "突出蓝色"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "雕刻"
         Index           =   3
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "氖化"
         Index           =   5
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "模糊 (少)"
         Index           =   7
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "模糊 (多)"
         Index           =   8
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "锐化"
         Index           =   10
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "添加杂色..."
         Index           =   12
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "腐蚀..."
         Index           =   14
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "煞风..."
         Index           =   15
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "添加雾道..."
         Index           =   16
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "添加图象噪音"
         Index           =   17
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "冻结 (少)"
         Index           =   19
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "冻结 (多)"
         Index           =   20
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "黑白化"
         Index           =   22
         Begin VB.Menu mnuBW 
            Caption         =   "模式 1"
            Index           =   0
         End
         Begin VB.Menu mnuBW 
            Caption         =   "模式 2"
            Index           =   1
         End
         Begin VB.Menu mnuBW 
            Caption         =   "模式 3"
            Index           =   2
         End
         Begin VB.Menu mnuBW 
            Caption         =   "灰度模式"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "软型染色"
         Index           =   24
         Begin VB.Menu mnuSoft 
            Caption         =   "红色"
            Index           =   0
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "绿色"
            Index           =   1
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "橙色"
            Index           =   2
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "黄色"
            Index           =   3
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "紫色"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "硬型染色"
         Index           =   26
         Begin VB.Menu mnuHard 
            Caption         =   "红色"
            Index           =   0
         End
         Begin VB.Menu mnuHard 
            Caption         =   "绿色"
            Index           =   1
         End
         Begin VB.Menu mnuHard 
            Caption         =   "蓝色"
            Index           =   2
         End
         Begin VB.Menu mnuHard 
            Caption         =   "黄色"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuSpecialFilters 
      Caption         =   "特殊滤镜(&I)"
      Begin VB.Menu mnuSpFil 
         Caption         =   "灰色"
         Index           =   0
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "茶色"
         Index           =   1
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "水底"
         Index           =   2
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "黄色"
         Index           =   3
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "木炭"
         Index           =   4
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "夜色"
         Index           =   5
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "日食"
         Index           =   6
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "紫色"
         Index           =   7
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "幽灵"
         Index           =   8
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "虚幻灰暗"
         Index           =   9
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "强化冷暖色"
         Index           =   10
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "惨烈"
         Index           =   11
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "斑点化"
         Index           =   12
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "暴光过度"
         Index           =   13
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "纸质杂色"
         Index           =   14
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "修饰(&W)"
      Begin VB.Menu mnuEff 
         Caption         =   "侧光水平百叶窗..."
         Index           =   0
      End
      Begin VB.Menu mnuEff 
         Caption         =   "侧光垂直百叶窗..."
         Index           =   1
      End
      Begin VB.Menu mnuEff 
         Caption         =   "平光水平百叶窗..."
         Index           =   2
      End
      Begin VB.Menu mnuEff 
         Caption         =   "平光垂直百叶窗..."
         Index           =   3
      End
      Begin VB.Menu mnuEff 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEff 
         Caption         =   "水平线"
         Index           =   5
      End
      Begin VB.Menu mnuEff 
         Caption         =   "垂直线"
         Index           =   6
      End
      Begin VB.Menu mnuEff 
         Caption         =   "正方形网状线"
         Index           =   7
      End
      Begin VB.Menu mnuEff 
         Caption         =   "方形层层深入线"
         Index           =   8
      End
      Begin VB.Menu mnuEff 
         Caption         =   "圆形层层深入线"
         Index           =   9
      End
      Begin VB.Menu mnuEff 
         Caption         =   "右斜线"
         Index           =   10
      End
      Begin VB.Menu mnuEff 
         Caption         =   "左斜线"
         Index           =   11
      End
      Begin VB.Menu mnuEff 
         Caption         =   "交叉斜线"
         Index           =   12
      End
      Begin VB.Menu mnuEff 
         Caption         =   "水平波浪线"
         Index           =   13
      End
      Begin VB.Menu mnuEff 
         Caption         =   "垂直波浪线"
         Index           =   14
      End
      Begin VB.Menu mnuEff 
         Caption         =   "向上水平波浪折线"
         Index           =   15
      End
      Begin VB.Menu mnuEff 
         Caption         =   "向左垂直波浪折线"
         Index           =   16
      End
      Begin VB.Menu mnuEff 
         Caption         =   "向下水平波浪折线"
         Index           =   17
      End
      Begin VB.Menu mnuEff 
         Caption         =   "向右垂直波浪折线"
         Index           =   18
      End
      Begin VB.Menu mnuEff 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuEff 
         Caption         =   "添加边框"
         Index           =   20
         Begin VB.Menu mnuBorder 
            Caption         =   "单色边框"
            Index           =   0
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "扩边单色边框"
            Index           =   1
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "渐变边框 1"
            Index           =   2
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "扩边渐变边框 1"
            Index           =   3
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "渐变边框 1"
            Index           =   4
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "扩边渐变边框 1"
            Index           =   5
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "圆形单色边框"
            Index           =   7
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "圆形渐变边框 1"
            Index           =   8
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "圆形渐变边框 2"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuMixing 
      Caption         =   "填充(&M)"
      Begin VB.Menu mnuMix 
         Caption         =   "单色填充"
         Index           =   0
      End
      Begin VB.Menu mnuMix 
         Caption         =   "渐变填充 1"
         Index           =   1
      End
      Begin VB.Menu mnuMix 
         Caption         =   "渐变填充 2"
         Index           =   2
      End
      Begin VB.Menu mnuMix 
         Caption         =   "方形填充 1"
         Index           =   3
      End
      Begin VB.Menu mnuMix 
         Caption         =   "方形填充 2"
         Index           =   4
      End
      Begin VB.Menu mnuMix 
         Caption         =   "圆形填充 1"
         Index           =   5
      End
      Begin VB.Menu mnuMix 
         Caption         =   "圆形填充 2"
         Index           =   6
      End
      Begin VB.Menu mnuMix 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuMix 
         Caption         =   "居中填充图片"
         Index           =   8
      End
      Begin VB.Menu mnuMix 
         Caption         =   "平铺填充图片"
         Index           =   9
      End
   End
   Begin VB.Menu mnuDeformation 
      Caption         =   "变形(&D)"
      Begin VB.Menu mnuDef 
         Caption         =   "层层递进"
         Index           =   0
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuDef 
         Caption         =   "马赛克"
         Index           =   2
      End
      Begin VB.Menu mnuDef 
         Caption         =   "圆形马赛克"
         Index           =   3
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuDef 
         Caption         =   "水平波浪处理"
         Index           =   5
      End
      Begin VB.Menu mnuDef 
         Caption         =   "反方向水平波浪处理"
         Index           =   6
      End
      Begin VB.Menu mnuDef 
         Caption         =   "垂直波浪处理"
         Index           =   7
      End
      Begin VB.Menu mnuDef 
         Caption         =   "反向垂直波浪处理"
         Index           =   8
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuDef 
         Caption         =   "分割图片"
         Index           =   10
      End
   End
   Begin VB.Menu mnuEdges 
      Caption         =   "渐变边(&L)"
      Begin VB.Menu mnuEdge 
         Caption         =   "增加渐变边 (少)"
         Index           =   0
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "增加渐变边 (多)"
         Index           =   1
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "自定义增加"
         Index           =   2
         Begin VB.Menu mnuPEdge 
            Caption         =   "左边 1"
            Index           =   0
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "左边 2"
            Index           =   1
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "左边 3"
            Index           =   2
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "右边 1"
            Index           =   4
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "右边 2"
            Index           =   5
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "右边 3"
            Index           =   6
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "上边 1"
            Index           =   8
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "上边 2"
            Index           =   9
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "上边 3"
            Index           =   10
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "底边 1"
            Index           =   12
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "底边 2"
            Index           =   13
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "底边 3"
            Index           =   14
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   15
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "左右 1"
            Index           =   16
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "左右 2"
            Index           =   17
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   18
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "上下1"
            Index           =   19
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "上下2"
            Index           =   20
         End
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "文字(&T)"
      Begin VB.Menu mnuTxt 
         Caption         =   "添加文字"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAb 
         Caption         =   "帮助文档..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuJ2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAb2 
         Caption         =   "关于..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "MnuTool"
      Visible         =   0   'False
      Begin VB.Menu mnuHistory1 
         Caption         =   "History"
      End
      Begin VB.Menu mnuSelection1 
         Caption         =   "Selection"
      End
      Begin VB.Menu mnuMagnifier 
         Caption         =   "Magnifier"
      End
      Begin VB.Menu jt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseMe 
         Caption         =   "Close me"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBack_Click()
ToolHide
End Sub

Private Sub CmdCloseTol_Click(Index As Integer)
ToolF(Index).Visible = False
End Sub

Private Sub Command1_Click()
FSetting.Show 1
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub

Private Sub Command12_Click()
Tol1(4).Visible = False
Tol1(5).Visible = True
Tol1(6).Visible = False
End Sub

Private Sub Command13_Click()
Tol1(4).Visible = True
Tol1(5).Visible = False
Tol1(6).Visible = False
End Sub

Private Sub Command14_Click()
Tol1(4).Visible = False
Tol1(5).Visible = False
Tol1(6).Visible = True
End Sub

Private Sub Command15_Click()
Tol1(4).Visible = False
Tol1(5).Visible = True
Tol1(6).Visible = False
End Sub

Private Sub Command18_Click()
Tol1(9).Visible = False
Tol1(10).Visible = True
End Sub

Private Sub Command19_Click()
Tol1(9).Visible = True
Tol1(10).Visible = False
End Sub

Private Sub Command2_Click(Index As Integer)
FCS.SetFocus
ToolShow  ' Show!
ToolUseImg.Picture = Image23(Index).Picture
ToolName.Caption = Command2(Index).Caption
ToolName2.Caption = ToolName.Caption


If mnuFilters.Enabled = True Then
'Add more!!
    If Index = 0 Then
    ToolFunc(0).Visible = True
    ElseIf Index = 1 Then
    ToolFunc(1).Visible = True
    ElseIf Index = 2 Then
    ToolFunc(2).Visible = True
    ElseIf Index = 3 Then
    ToolFunc(3).Visible = True
    ElseIf Index = 4 Then
    ToolFunc(4).Visible = True
    ElseIf Index = 5 Then
    ToolFunc(6).Visible = True
    ElseIf Index = 6 Then
    ToolFunc(5).Visible = True
    Else
CmdBack_Click
End If

End If

End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
FCS.SetFocus
End Sub


Private Sub Command4_Click()
Tol1(0).Visible = False
Tol1(1).Visible = True
End Sub

Private Sub Command5_Click()
Tol1(0).Visible = True
Tol1(1).Visible = False
End Sub

Private Sub Command8_Click()
Tol1(2).Visible = False
Tol1(3).Visible = True
End Sub

Private Sub Command9_Click()
Tol1(2).Visible = True
Tol1(3).Visible = False
End Sub




Private Sub HS1_Scroll()
Pic1.Left = HS1.Value
Pic1.SetFocus
End Sub

Private Sub IBd1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IBd2.Visible = True
End Sub

Private Sub IBd2_Click()

HideTip
IBd2.Visible = False
IBd3.Visible = True
IBg3.Visible = False
ShowTool = True
ToolXY.Visible = True
ToolRedo.Visible = True
ToolZoom.Visible = True
FrDS.Visible = True

End Sub

Private Sub IBd3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IBg2.Visible = False
End Sub

Private Sub IBg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IBg2.Visible = True
End Sub

Private Sub IBg2_Click()
IBg2.Visible = False
IBg3.Visible = True
IBd3.Visible = False
ShowTool = False
ToolXY.Visible = False
ToolRedo.Visible = False
ToolZoom.Visible = False
FrDS.Visible = False
End Sub

Private Sub IBg3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IBd2.Visible = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = False
Bover.Visible = False
End Sub





Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = False
Bover.Visible = False
IBg2.Visible = False
IBd2.Visible = False
FCS.SetFocus
On Error Resume Next
    For sti = 0 To 100
    SelectTool(sti).BackStyle = 0
    Next
End Sub



Private Sub Image5_Click()
HideTip

End Sub

Private Sub ImgT1_Click(Index As Integer)
If Index <> 0 Then  'except the ORIGINAL
SelectTool_Click Index
End If
End Sub

Private Sub ImgT1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> 0 Then  'except the ORIGINAL
SelectTool_MouseMove Index, 0, 0, 1, 1
End If
End Sub

Private Sub ImgT2_Click(Index As Integer)
Select Case Index
Case 0
mnuEmbossSpecial_Click 0
Case 1
mnuEmbossSpecial_Click 1
Case 2
mnuEmbossSpecial_Click 2
Case 3
mnuSoft_Click 0
Case 4
mnuSoft_Click 1
Case 5
mnuSoft_Click 2
Case 6
mnuSoft_Click 3
Case 7
mnuSoft_Click 4

End Select
End Sub

Private Sub ImgT2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For oe = 0 To 20
ImgT2(oe).Borderstyle = 0
Next
ImgT2(Index).Borderstyle = 1

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


Sub Reload1()
HS1.Enabled = False
VS1.Enabled = False
PB1.Value = 0
Toolbar1.Buttons(1).Enabled = False
Cancel1.Enabled = False
ClearMem
mnuColors.Enabled = False
mnuPicture.Enabled = False
mnuSelection.Enabled = False
mnuFiles(1).Enabled = False
mnuFiles(2).Enabled = False
mnuFiles(4).Enabled = False
mnuFilters.Enabled = False
mnuSpecialFilters.Enabled = False
mnuEffects.Enabled = False
mnuMixing.Enabled = False
mnuDeformation.Enabled = False
mnuEdges.Enabled = False
mnuText.Enabled = False
Label4.Caption = ""
ToolXY.Label4.Caption = ""

Label1.Caption = ""
        PicX.Visible = False
        Run2.Visible = True
        Inject1.Visible = True
        For pr = 0 To 4
        Image1(pr).Refresh
        Image1(pr).Enabled = False
        Next
        For Lr = 0 To 5
        Label8(Lr).Enabled = False
        Next
End Sub
Private Sub Cancel1_Click()
Redo
End Sub


Private Sub Form_Load()

On Error Resume Next


LoadLang1


apbg = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "APbg")
apbgadd = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "APbgadd")
If apbg = "" Then ' Default
Me.Picture = LoadPicture(App.Path & "\APbg.skin")
ElseIf apbg = "0" Then ' None

ElseIf apbg = "1" Then ' Custom
Me.Picture = LoadPicture(apbgadd)
End If


'MkDir App.Path & "C:\Sica Pictures\"
'File1.Path = App.Path & "C:\Sica Pictures\"
Set Shape1.Container = Pic1
Scol(0) = &H202020
Scol(1) = &H404040
Scol(2) = &H606060
Scol(3) = &H808080
Scol(4) = &HA0A0A0
Scol(5) = &HC0C0C0
Scol(6) = &HE0E0E0
Scol(7) = &HFFFFFF
Scol(8) = &HE0E0E0
Scol(9) = &HC0C0C0
Scol(10) = &HA0A0A0
Scol(11) = &H808080
Scol(12) = &H606060
Scol(13) = &H404040
Scol(14) = &H202020
Scol(15) = 0
'Moves!!!
PicX.Move 300, 26, 487, 487
FrDS.Move 8, 24, 272, 497
'ToolF(0).Move 360, 1800, 3255, 4575
Tol1(0).Move 120, 1680, 3615, 6465
Tol1(1).Move 120, 1680, 3615, 6465
ToolF(0).Move 360, 2280, 3255, 4575
ToolF(1).Move 360, 1560, 3255, 6255
For TF1 = 0 To 6  '排列面板
ToolFunc(TF1).Move 120, 1320
Next
Tol1(0).Visible = True
Tol1(1).Visible = False

HS1.Enabled = False
VS1.Enabled = False
PB1.Value = 0
Toolbar1.Buttons(1).Enabled = False
Cancel1.Enabled = False
ClearMem
mnuColors.Enabled = False
mnuPicture.Enabled = False
mnuSelection.Enabled = False
mnuFiles(1).Enabled = False
mnuFiles(2).Enabled = False
mnuFiles(4).Enabled = False
mnuFilters.Enabled = False
mnuSpecialFilters.Enabled = False
mnuEffects.Enabled = False
mnuMixing.Enabled = False
mnuDeformation.Enabled = False
mnuEdges.Enabled = False
mnuText.Enabled = False
ShowTool = False
SelectAll
For Xx = 0 To 1
Set Text(Xx).Container = Pic1
Text(Xx).BackStyle = 0
Next Xx
EnumFonts Printer.hDC, vbNullString, AddressOf EnumFontProc, 0


 FLoading2.Hide
 Unload FLoading2

Timer4.Enabled = True  ' Tip's timer

Dim regWay As String
regWay = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Way")
If regWay = "1" Then
IBd2_Click
End If
End Sub


Sub LoadLang1()
On Error Resume Next

'Lang Selection
If LangA = "lge" Then
mnuSelLang(0).Checked = True
mnuSelLang(1).Checked = False
ElseIf LangA = "lgc" Then
mnuSelLang(0).Checked = False
mnuSelLang(1).Checked = True
End If



FTitle = lgT(314) 'Caption Load


Timer3.Enabled = True
 munReg1.Visible = False
 Caption = FTitle
 Inject1.Caption = lgT(400) 'Introduce






'Lang     ===============1
mnuFile.Caption = lgT(10)
'----------------------------
For m1 = 0 To 8
mnuFiles(m1).Caption = lgT(m1 + 22)
Next
For m1a = 0 To 1
mnuSelLang(m1a).Caption = lgT(m1a + 31)
Next
'==============================2
mnuSelection.Caption = lgT(11)
'----------------------------
Cancel1.Caption = lgT(41)
For m2 = 0 To 2
mnuSel(m2).Caption = lgT(m2 + 43)
Next
'==============================3
mnuPicture.Caption = lgT(12)
'----------------------------
For m3 = 0 To 6
mnuPic(m3).Caption = lgT(m3 + 56)
Next
'==============================4
mnuColors.Caption = lgT(13)
'----------------------------
For m4 = 0 To 13
mnuCol(m4).Caption = lgT(m4 + 73)
Next
For m41 = 0 To 2
mnuKill(m41).Caption = lgT(m41 + 100)
Next
For m43 = 0 To 2
mnuInv(m43).Caption = lgT(m43 + 103)
Next

'==============================5
mnuFilters.Caption = lgT(14)
'----------------------------
For m5 = 0 To 26
mnuFilter(m5).Caption = lgT(m5 + 106)
Next

For m51 = 0 To 2
mnuEmbossSpecial(m51).Caption = lgT(m51 + 139)
Next

For m52 = 0 To 2
mnuBW(m52).Caption = lgT(m52 + 142)
Next
mnuBW(3).Caption = lgT(169)
'Grey


For m53 = 0 To 4
mnuSoft(m53).Caption = lgT(m53 + 145)
Next

For m54 = 0 To 3
mnuHard(m54).Caption = lgT(m54 + 150)
Next




'==============================6
mnuSpecialFilters.Caption = lgT(15)
'----------------------------
For m6 = 0 To 14
mnuSpFil(m6).Caption = lgT(m6 + 154)
Next

'==============================7
mnuEffects.Caption = lgT(16)
'----------------------------
For m7 = 0 To 20
mnuEff(m7).Caption = lgT(m7 + 180)
Next
For m71 = 0 To 9
mnuBorder(m71).Caption = lgT(m71 + 210)
Next

'==============================8
mnuMixing.Caption = lgT(17)
'----------------------------
For m8 = 0 To 9
mnuMix(m8).Caption = lgT(m8 + 220)
Next
'==============================9
mnuDeformation.Caption = lgT(18)
'----------------------------
For m9 = 0 To 10
mnuDef(m9).Caption = lgT(m9 + 240)
Next
'==============================10
mnuEdges.Caption = lgT(19)
'----------------------------
For k10 = 0 To 2
mnuEdge(k10).Caption = lgT(k10 + 251)
Next
For k101 = 0 To 20
mnuPEdge(k101).Caption = lgT(k101 + 260)
Next

'==============================11
mnuText.Caption = lgT(20)
'----------------------------
mnuTxt(0).Caption = lgT(281)
'==============================12
mnuAbout.Caption = lgT(21)
'----------------------------
mnuAb.Caption = lgT(286)
munReg1.Caption = lgT(376)
munWeb1.Caption = lgT(377)
mnuAb2.Caption = lgT(288)

'==============================
Label3.Caption = lgT(295) 'Func. Text
ToolRedo.Label3.Caption = Label3.Caption
Label5.Caption = lgT(296)
Label6.Caption = lgT(297)
ToolXY.Label3.Caption = Label6.Caption
Label7.Caption = lgT(298)
ToolZoom.Label3.Caption = lgT(294)
'
For shortout1 = 0 To 5
Label8(shortout1).Caption = lgT(shortout1 + 299)
Next
'
Run1.Caption = lgT(328)
Run2.Caption = lgT(329)

 ' 补充类 语言
If LangA = "lgc" Then
mnuPlugins.Caption = "插件(&P)"
Command1.Caption = "参数设置(&S)..."
mnuUpdate.Caption = "检查更新"
LabNote1.Caption = "提示：已发现可用更新，点这里查看详细信息..."
mnuSup1.Caption = "服务信息"
Label10.Caption = "画室"
LabT1(0).Caption = "原图"
mnuHistory1.Caption = "历史操作"
mnuSelection1.Caption = "选择区域"
mnuMagnifier.Caption = "放大辅助"
mnuCloseMe.Caption = "关闭本栏"
On Error Resume Next
For i = 0 To 15
LabTi(i).Caption = LabT1(0).Caption
Next

Command3.Caption = "<< 上页"
Command4.Caption = "下页 >>"
Command5.Caption = Command3.Caption
Command6.Caption = Command4.Caption
Command7.Caption = Command3.Caption
Command8.Caption = Command4.Caption
Command9.Caption = Command3.Caption
Command10.Caption = Command4.Caption

Command11.Caption = Command3.Caption
Command12.Caption = Command4.Caption
Command13.Caption = Command3.Caption
Command14.Caption = Command4.Caption
Command15.Caption = Command3.Caption
Command16.Caption = Command4.Caption
Command17.Caption = Command3.Caption
Command18.Caption = Command4.Caption
Command19.Caption = Command3.Caption
Command20.Caption = Command4.Caption


Else
mnuPlugins.Caption = "&Plugin"

mnuUpdate.Caption = "Check Updates"
LabNote1.Caption = "Notice: Found available updates. Click here for details..."
mnuSup1.Caption = "Support"

Command2(0).Caption = lgT(14)
Command2(1).Caption = lgT(15)
Command2(2).Caption = lgT(16)
Command2(3).Caption = lgT(18)
Command2(4).Caption = lgT(12)
Command2(5).Caption = lgT(13)
Command2(6).Caption = lgT(17)
CmdBack.Caption = "<< Back"

Label9.Caption = "More intuitionistic!"
Label11.Caption = "Click here to experience the intuitionistic operations"
End If


LabT1(1).Caption = mnuFilter(0).Caption
LabT1(2).Caption = mnuFilter(3).Caption
LabT1(3).Caption = mnuFilter(1).Caption
LabT1(4).Caption = mnuFilter(5).Caption
LabT1(5).Caption = mnuFilter(7).Caption
LabT1(6).Caption = mnuFilter(10).Caption
LabT1(7).Caption = mnuFilter(12).Caption
LabT1(8).Caption = mnuFilter(22).Caption
LabT1(9).Caption = mnuFilter(14).Caption
LabT1(10).Caption = mnuFilter(15).Caption
LabT1(11).Caption = mnuFilter(16).Caption
LabT1(12).Caption = mnuFilter(17).Caption
LabT1(13).Caption = mnuFilter(19).Caption
LabT1(14).Caption = mnuFilter(24).Caption

'Special Filter
For sfe1 = 0 To 14
LabT1(sfe1 + 15).Caption = mnuSpFil(sfe1).Caption
Next
''''''''''

ToolF(0).Caption = mnuFilter(1).Caption
ToolF(1).Caption = mnuFilter(24).Caption
' soft color AND hard color
For eef1 = 0 To 2
Leffect(eef1).Caption = mnuEmbossSpecial(eef1).Caption
Next
For eef2 = 0 To 4
Leffect(eef2 + 3).Caption = mnuSoft(eef2).Caption
Next

' Effects
For eFl = 0 To 3
LabT1(eFl + 30).Caption = mnuEff(eFl).Caption
Next

For efl2 = 0 To 13
LabT1(efl2 + 34).Caption = mnuEff(efl2 + 5).Caption
Next

LabT1(48).Caption = mnuEff(20).Caption
' Picture
For pc1 = 0 To 1
LabT1(49 + pc1).Caption = mnuPic(pc1).Caption
Next
For pc2 = 0 To 3
LabT1(51 + pc2).Caption = mnuPic(pc2 + 3).Caption
Next

' Mixing
For Mc1 = 0 To 6
LabT1(63 + Mc1).Caption = mnuMix(Mc1).Caption
Next
LabT1(70).Caption = mnuMix(8).Caption
LabT1(71).Caption = mnuMix(9).Caption

' Def
LabT1(55).Caption = mnuDef(0).Caption
LabT1(56).Caption = mnuDef(2).Caption
LabT1(57).Caption = mnuDef(3).Caption
LabT1(58).Caption = mnuDef(5).Caption
LabT1(59).Caption = mnuDef(6).Caption
LabT1(60).Caption = mnuDef(7).Caption
LabT1(61).Caption = mnuDef(8).Caption
LabT1(62).Caption = mnuDef(10).Caption

' Color


LabT1(72).Caption = mnuCol(0).Caption
LabT1(73).Caption = mnuCol(2).Caption
LabT1(74).Caption = mnuCol(4).Caption
LabT1(79).Caption = mnuCol(11).Caption

LabT1(75).Caption = mnuCol(6).Caption
LabT1(76).Caption = mnuCol(7).Caption
LabT1(77).Caption = mnuCol(9).Caption
LabT1(78).Caption = mnuCol(13).Caption





 FLoading2.Hide
 Unload FLoading2
 Unload FLogo
 
FMain.WindowState = 2

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    For sti = 0 To 100
    SelectTool(sti).BackStyle = 0
    Next
For i = 0 To 5
Label8(i).ForeColor = &H404080
Next
Bmove.Visible = False
Bover.Visible = False
IBg2.Visible = False
IBd2.Visible = False
FCS.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
On Error Resume Next
    If mnuFiles(1).Enabled = True Then
'FIXIT: Declare 'e' with an early-bound data type                                          FixIT90210ae-R1672-R1B8ZE
    Dim e
    e = MsgBox(lgT(0), vbExclamation + vbYesNoCancel, "Exit") '您要在退出前保存图象吗
        If e = vbYes Then
        CD1.Filter = "Bitmap(.bmp)|*.bmp"
        CD1.flags = 2
        CD1.FileName = 0
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
        PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
        CD1.FileName = PicFileName
        CD1.ShowSave
        PicFileName = CD1.FileTitle
        SavePicture Pic1.Image, PicFileName
        SetPicInfo
        DoEvents
        qag = MsgBox(lgT(399), vbExclamation + vbYesNo, "End Process")
            If lgT(381) <> "YO" Then ' If reged
            Fagm.Show 1
            End If
        If qag = vbYes Then Cancel = False        '!!!!!!!!!!
        ElseIf e = vbNo Then
     Cancel = False
         End If
    Else
  Cancel = False
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next

'过小强制保持大小
If Me.Height < 5850 Then
Me.Height = 5850
End If
If Me.Width < 10600 Then
Me.Width = 10600
End If
''''''''''''''''''
FrDS.Height = Me.Height / 15 - 75
PicX.Width = Me.Width / 15 - 308
PicX.Height = Me.Height / 15 - 78
PB1.Top = Me.Height / 15 - 48
SetScrollBars
End Sub

Private Sub Form_Unload(Cancel As Integer)
FWhole.mAP.Enabled = True
FWhole.mAP.Checked = False
If FWhole.mIC.Checked = False Then
On Error Resume Next
ToolsClose
 FWel.Show
 FWel.SetFocus
 Unload frmTray
  ' Load Tool List
End If
End Sub

Private Sub HelpImg1_Click()
Fhelp1.Show 1
End Sub

Private Sub HelpImg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
HelpImg1.Borderstyle = 1
End Sub

Private Sub HelpImg1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
HelpImg1.Borderstyle = 0
End Sub

Private Sub HS1_GotFocus()
Dummy.SetFocus
End Sub
Private Sub Image1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
Redo
ElseIf Index = 1 Then
Redo
Redo
ElseIf Index = 2 Then
Redo
Redo
Redo
ElseIf Index = 3 Then
Redo
Redo
Redo
Redo
ElseIf Index = 4 Then
Redo
Redo
Redo
Redo
Redo
End If
End Sub



Private Sub Label11_Click()
HideTip
IBd2_Click
End Sub

Private Sub Label2_Change()
SBar1.Panels(1).Text = Label2.Caption
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = False
Bover.Visible = False
End Sub

Private Sub Label8_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
PopupMenu Me.mnuColors, , (Label8(Index).Left + Label8(Index).Width - 80), (Label8(Index).Top)
Case 1
PopupMenu Me.mnuEffects, , (Label8(Index).Left + Label8(Index).Width - 80), (Label8(Index).Top)
Case 2
PopupMenu Me.mnuDeformation, , (Label8(Index).Left + Label8(Index).Width - 80), (Label8(Index).Top)
Case 3
FText.Show 1
Case 4
PopupMenu Me.mnuPicture, , (Label8(Index).Left + Label8(Index).Width - 80), (Label8(Index).Top)
Case 5
PopupMenu Me.mnuFilters, , (Label8(Index).Left + Label8(Index).Width - 80), (Label8(Index).Top)
End Select
Label8(Index).ForeColor = &H404080
End Sub

Private Sub Label8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For i = 0 To 5
Label8(i).ForeColor = &H404080
Next
Label8(Index).ForeColor = &H80FF&
End Sub

Private Sub Label9_Click()
HideTip
IBd2_Click
End Sub

Private Sub LabNote1_Click()
mnuUpdate_Click
End Sub

Private Sub LabT1_Click(Index As Integer)
ImgT1_Click Index
End Sub

Private Sub LabT1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgT1_MouseMove Index, 0, 0, 1, 1
End Sub

Private Sub Leffect_Click(Index As Integer)
ImgT2_Click Index
End Sub

Private Sub Leffect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgT2_MouseMove Index, 0, 0, 1, 1

End Sub

Private Sub mnuAb_Click()
On Error Resume Next
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Help\ap.htm", "", App.Path, 1
End Sub

Private Sub mnuAb2_Click()
about1.Show
End Sub

Private Sub mnuBorder_Click(Index As Integer)
Select Case Index
Case 0 'solid border
Col = 26
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 1 'solid border reduced
Col = 27
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 2 'gradient border 1
Col = 28
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 3 'gradient border 1 reduced
Col = 29
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 4 'gradient border 2
Col = 30
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 5 'gradient border 2 reduced
Col = 31
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 7 'solid circular border
Col = 32
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 8 'gradient circular border 1
Col = 33
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 9 'gradient circular border 1
Col = 34
FColor.Caption = "Effect - Borders"
FColor.Show 1
End Select
End Sub

Private Sub mnuBW_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
BnW Xcor0, Ycor0, Xcor1, Ycor1, 200
Case 1
BnW Xcor0, Ycor0, Xcor1, Ycor1, 150
Case 2
BnW Xcor0, Ycor0, Xcor1, Ycor1, 100
Case 3
mnuCol_Click 13
End Select
End Sub

Private Sub mnuCloseMe_Click()
On Error GoTo Ex1
If ActiveTool = 0 Then
ToolRedo.Visible = False
mnuHistory1.Checked = False
ElseIf ActiveTool = 1 Then
ToolXY.Visible = False
mnuSelection1.Checked = False
ElseIf ActiveTool = 2 Then
ToolZoom.Visible = False
mnuMagnifier.Checked = False
End If
Exit Sub
Ex1:
MsgBox "Error", vbInformation, "Error"
End Sub

Private Sub mnuCol_Click(Index As Integer)
Select Case Index
Case 0 'color comp
Col = 0
FColor.Caption = lgT(315)
FColor.Show 1
Case 6 'brighten
Col = 1
FColor.Caption = lgT(315)
FColor.Show 1
Case 7 'contrast
Col = 3
FColor.Caption = lgT(315)
FColor.Show 1
Case 9 'fotoneg.
SaveRedo
PhotoNeg Xcor0, Ycor0, Xcor1, Ycor1
Case 13 'grey
SaveRedo
GreyColor Xcor0, Ycor0, Xcor1, Ycor1
End Select
End Sub

Private Sub mnuDef_Click(Index As Integer)
FColor.Caption = lgT(316)
Select Case Index
Case 0
FEcho.Show 1
Case 2 'mozaic
Col = 42
FColor.Show 1
Case 3 'blurred mozaic
Col = 43
FColor.Show 1
Case 5 'wave X
Col = 44
FColor.Show 1
Case 6 'abs wave X
Col = 45
FColor.Show 1
Case 7 'wave Y
Col = 46
FColor.Show 1
Case 8 'abs wave Y
Col = 47
FColor.Show 1
Case 10 'tile pic
Col = 49
FColor.Show 1
End Select
End Sub

Private Sub mnuEdge_Click(Index As Integer)
Select Case Index
Case 0 'inc edges
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 1 'more inc edges
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
End Select
End Sub

Private Sub mnuEff_Click(Index As Integer)
Select Case Index
Case 0 'H blinds
Col = 8
FColor.Caption = lgT(317)
FColor.Show 1
Case 1 'V blinds
Col = 9
FColor.Caption = lgT(317)
FColor.Show 1
Case 2 ' bump H blinds
Col = 10
FColor.Caption = lgT(317)
FColor.Show 1
Case 3 ' bump V blinds
Col = 11
FColor.Caption = lgT(317)
FColor.Show 1
Case 5 ' add hor lines
Col = 12
FColor.Caption = lgT(317)
FColor.Show 1
Case 6 ' add ver lines
Col = 13
FColor.Caption = lgT(317)
FColor.Show 1
Case 7 ' add squares
Col = 14
FColor.Caption = lgT(317)
FColor.Show 1
Case 8 ' add squares
Col = 15
FColor.Caption = lgT(317)
FColor.Show 1
Case 9 ' add squares
Col = 16
FColor.Caption = lgT(317)
FColor.Show 1
Case 10 ' add dia R lines
Col = 17
FColor.Caption = lgT(317)
FColor.Show 1
Case 11 ' add dia R lines
Col = 18
FColor.Caption = lgT(317)
FColor.Show 1
Case 12 ' add crossed lines
Col = 19
FColor.Caption = lgT(317)
FColor.Show 1
Case 13 ' add H wave lines
Col = 20
FColor.Caption = lgT(317)
FColor.Show 1
Case 14 ' add V wave lines
Col = 21
FColor.Caption = lgT(317)
FColor.Show 1
Case 15 ' add abs H wave lines
Col = 22
FColor.Caption = lgT(317)
FColor.Show 1
Case 16 ' add abs V wave lines
Col = 23
FColor.Caption = lgT(317)
FColor.Show 1
Case 17 ' add abs H wave lines reversed
Col = 24
FColor.Caption = lgT(317)
FColor.Show 1
Case 18 ' add abs V wave lines reversed
Col = 25
FColor.Caption = lgT(317)
FColor.Show 1
End Select
End Sub

Private Sub mnuEmbossSpecial_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0 'holdred
HoldRed Xcor0, Ycor0, Xcor1, Ycor1
Case 1 'holdgreen
HoldGreen Xcor0, Ycor0, Xcor1, Ycor1
Case 2 'holdblue
HoldBlue Xcor0, Ycor0, Xcor1, Ycor1
End Select
End Sub


Sub OpenFile1()
mnuFiles_Click (0)
End Sub




Private Sub mnuFiles_Click(Index As Integer)
On Error GoTo ExitIt
Select Case Index
                                            Case 0 'open pic
    If mnuFiles(1).Enabled = True Then
'FIXIT: Declare 'a' with an early-bound data type                                          FixIT90210ae-R1672-R1B8ZE
    Dim a
    a = MsgBox(lgT(1), vbInformation + vbYesNoCancel, "Save") '"您要保存编辑中的图片吗？
        If a = vbYes Then
        CD1.Filter = "Bitmap(.bmp)|*.bmp"
        CD1.flags = 2
        CD1.FileName = 0
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
        PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
        CD1.FileName = PicFileName
        CD1.ShowSave
        PicFileName = CD1.FileTitle
        SavePicture Pic1.Image, PicFileName
        SetPicInfo
        DoEvents
        ElseIf a = vbCancel Then
        Exit Sub
        End If
    End If
'cd1.Filter = lgT(3) & "|*.bmp;*.jpg;*.gif;*.wmf;*.ico"
'cd1.flags = 2
    'If Run1.Visible = False Then
    SProject.Show 1
        'cd1.ShowOpen
    'Else
        On Error Resume Next
                If Dir("Temp.ini") = "" Then
                    Exit Sub
                Else
                    Dim load1 As String
                    Open "Temp.ini" For Input As #1
                    Do While Not EOF(1)
                    Input #1, load1
                    Loop
                    Close
                    CD1.FileName = load1
                    CD1.FileTitle = load1
                    Kill ("Temp.ini")
                End If

    'End If
'MsgBox CD1.FileName
'MsgBox CD1.FileTitle
FLoading.Show
Timer2.Enabled = True
Pic1.Picture = LoadPicture(CD1.FileName)
PicFileName = CD1.FileTitle
SetScrollBars
SetPicInfo
DoEvents
SelectAll
MemCount = 0
Toolbar1.Buttons(1).Enabled = False
Cancel1.Enabled = False
ClearMem
mnuColors.Enabled = True
mnuPicture.Enabled = True
mnuSelection.Enabled = True
mnuFiles(1).Enabled = True
mnuFiles(2).Enabled = True
mnuFiles(4).Enabled = True
mnuFilters.Enabled = True
mnuSpecialFilters.Enabled = True
mnuEffects.Enabled = True
mnuMixing.Enabled = True
mnuDeformation.Enabled = True
mnuEdges.Enabled = True
mnuText.Enabled = True
Run1.Visible = False
Run2.Visible = False
Inject1.Visible = False
Command1.Visible = False
For rx = 0 To 4
    Image1(rx).Enabled = False
Next
Unload FLoading
Timer2.Enabled = False
For i = 0 To 5
    Label8(i).Enabled = True
Next
PicX.Visible = True
                                        Case 1 'save picture
CD1.Filter = "Bitmap(.bmp)|*.bmp"
CD1.flags = 2
CD1.FileName = 0
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
CD1.FileName = PicFileName
CD1.ShowSave
PicFileName = CD1.FileTitle
SavePicture Pic1.Image, PicFileName
SetPicInfo
DoEvents
                                        Case 2 'save to spec. map"
'FIXIT: Replace 'Right' function with 'Right$' function                                    FixIT90210ae-R9757-R1B8ZE
If LCase(Right(PicFileName, 3)) <> "bmp" Then
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
End If
SavePicture Pic1.Image, App.Path & "C:\Sica Pictures\" & PicFileName
File1.Refresh
CD1.FileTitle = PicFileName
SetPicInfo
DoEvents
MsgBox "已保存至“我的图片”(C:\Sica Pictures\)", vbInformation, "保存完毕"
                                  
                                                Case 4 'print
MsgBox lgT(319), vbInformation, "Print"
Temp = MsgBox(lgT(320), vbQuestion + vbYesNoCancel, FTitle)
If Temp = vbCancel Then MsgBox lgT(321), , FTitle: Exit Sub
If Temp = vbYes Then Printer.Orientation = 2
If Temp = vbNo Then Printer.Orientation = 1
Printer.PaintPicture Pic1.Image, 0, 0
Printer.EndDoc
Printer.Orientation = 1
                                                    Case 5
    If mnuFiles(1).Enabled = True Then
'FIXIT: Declare 'C' with an early-bound data type                                          FixIT90210ae-R1672-R1B8ZE
    Dim C
    C = MsgBox(lgT(1), vbInformation + vbYesNoCancel, "Save") '"您要保存编辑中的图片吗？
        If C = vbYes Then
        CD1.Filter = "Bitmap(.bmp)|*.bmp"
        CD1.flags = 2
        CD1.FileName = 0
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
        PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
        CD1.FileName = PicFileName
        CD1.ShowSave
        PicFileName = CD1.FileTitle
        SavePicture Pic1.Image, PicFileName
        SetPicInfo
        DoEvents
        Reload1
        SProject.Show 1
        ElseIf C = vbNo Then
        Reload1
        SProject.Show 1
        ElseIf C = vbCancel Then
        Exit Sub
        End If
    Else
        Reload1
        SProject.Show 1
    End If
                                                Case 7
    
    
    
                                                Case 8
Unload Me

End Select
Exit Sub
ExitIt:
If Err.Number = 32755 Then
Exit Sub 'user pressed cancel
Else
Resume Next
End If
'MsgBox "Error # " & Err.Number & " - " & Err.Description, vbInformation
End Sub

Private Sub mnuFilter_Click(Index As Integer)
Select Case Index
Case 0 'emboss
SaveRedo
EmbossPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 3 'engrave
SaveRedo
EngravePicture Xcor0, Ycor0, Xcor1, Ycor1
Case 5 'neon
SaveRedo
NeonPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 7 'blur
SaveRedo
BlurPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 8 'blur more
SaveRedo
BlurPictureMore Xcor0, Ycor0, Xcor1, Ycor1
Case 10 'sharpen
SaveRedo
SharpenPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 12 'diffuse
Col = 4
FColor.Caption = lgT(322)
FColor.Show 1
Case 14 'erode
Col = 5
FColor.Caption = lgT(322)
FColor.Show 1
Case 15 'Blow
Col = 6
FColor.Caption = lgT(322)
FColor.Show 1
Case 16 'fog
Col = 7
FColor.Caption = lgT(322)
FColor.Show 1
Case 17 'noise
SaveRedo
AddNoise Xcor0, Ycor0, Xcor1, Ycor1
Case 19 'freeze
SaveRedo
FreezePic Xcor0, Ycor0, Xcor1, Ycor1, 1.1
Case 20 'freezemore
SaveRedo
FreezePic Xcor0, Ycor0, Xcor1, Ycor1, 1.5
End Select
End Sub

Private Sub mnuHistory1_Click()
ToolRedo.Show
End Sub

Private Sub mnuInv_Click(Index As Integer)
SaveRedo
PhotoNegComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnuKill_Click(Index As Integer)
SaveRedo
KillComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnumagnifier_Click()
ToolZoom.Show
End Sub

Private Sub mnuMix_Click(Index As Integer)
Select Case Index
Case 0 'mix solid color
Col = 35
FColor.Caption = lgT(323)
FColor.Show 1
Case 1 'mix gradient 1
Col = 36
FColor.Caption = lgT(323)
FColor.Show 1
Case 2 'mix gradient 2
Col = 37
FColor.Caption = lgT(323)
FColor.Show 1
Case 3 'mix box gradient 1
Col = 38
FColor.Caption = lgT(323)
FColor.Show 1
Case 4 'mix box gradient 1
Col = 39
FColor.Caption = lgT(323)
FColor.Show 1
Case 5 'mix circular gradient 1
Col = 40
FColor.Caption = lgT(323)
FColor.Show 1
Case 6 'mix circular gradient 2
Col = 41
FColor.Caption = lgT(323)
FColor.Show 1
'---------------
Case 8 'mix picture
Mix = 0
FPicture.Caption = lgT(324)
FPicture.Show 1
Case 9 'mix pattern
Mix = 1
FPicture.Caption = lgT(324)
FPicture.Show 1
End Select
End Sub

Private Sub mnuPEdge_Click(Index As Integer)
Select Case Index
Case 0 'inc edges L1
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
Case 1 'inc edges L2
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
Case 2 'inc edges L3
SaveRedo
BeginProcess
KillColXGrad3 0, 0, Pic1.Width, Pic1.Height
Case 4 'inc edges R1
SaveRedo
BeginProcess
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 5 'inc edges R2
SaveRedo
BeginProcess
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 6 'inc edges R3
SaveRedo
BeginProcess
KillColXGradRev3 0, 0, Pic1.Width, Pic1.Height
Case 8 'inc edges T1
SaveRedo
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
Case 9 'inc edges T2
SaveRedo
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
Case 10 'inc edges T3
SaveRedo
BeginProcess
KillColYGrad3 0, 0, Pic1.Width, Pic1.Height
Case 12 'inc edges B1
SaveRedo
BeginProcess
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 13 'inc edges B2
SaveRedo
BeginProcess
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 14 'inc edges B3
SaveRedo
BeginProcess
KillColYGradRev3 0, 0, Pic1.Width, Pic1.Height
Case 16 'inc edges L1 & R1
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 17 'inc edges L2& R2
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 19 'inc edges T1 & B1
SaveRedo
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 20 'inc edges T2& B2
SaveRedo
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
End Select
End Sub

Private Sub mnuPic_Click(Index As Integer)
Select Case Index
Case 0 ' flip X
SaveRedo
FlipX
Case 1 'flip Y
SaveRedo
FlipY
Case 3 'Mirror X
SaveRedo
MirrorX
Case 4 'Mirror Y
SaveRedo
MirrorXRev
Case 5 'Mirror X
SaveRedo
MirrorY
Case 6 'Mirror Y
SaveRedo
MirrorYRev
End Select
End Sub

Private Sub mnuPlugin_Click(Index As Integer)
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "Cannot Find Plugins Application", vbExclamation, "Error"
Else
ShellExecute Me.hWnd, "Open", "plugin.exe", "cs", App.Path, 1
End If
End Sub

Private Sub mnuPluginsAdd_Click()
With FAddPT
.Caption = "Add Plugins"
.Frame1(0).Visible = True
.Show 1
End With
End Sub

Private Sub mnuSel_Click(Index As Integer)
Select Case Index
Case 0 'adjust selection
If Shape1.Visible = False Then
MsgBox lgT(325), vbInformation, FTitle
Exit Sub
End If
Col = 48
FColor.Show 1
Case 1 'no selection
SelectAll

Case 2 'select all

Shape1.Visible = True
Xcor0 = 0: Ycor0 = 0
Xcor1 = 0: Ycor1 = 0
XXX1 = Xcor0: YYY1 = Ycor0
XXX2 = Xcor1: YYY2 = Ycor1
Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
SetCoordinates
FMain.Toolbar1.Buttons(3).Enabled = True
FMain.mnuSel(1).Enabled = True

Pic1.MousePointer = 2

'------------

            XXX2 = Pic1.Width - XXX1: YYY2 = Pic1.Height - YYY1
            Xcor0 = XXX1: Ycor0 = YYY1
            Xcor1 = XXX2: Ycor1 = YYY2
    If XXX2 < 0 Then
    Xcor0 = XXX1 + XXX2
            If Xcor0 < 0 Then Xcor0 = 0
    Xcor1 = XXX1 - Xcor0
    End If
    If YYY2 < 0 Then
    Ycor0 = YYY1 + YYY2
            If Ycor0 < 0 Then Ycor0 = 0
    Ycor1 = YYY1 - Ycor0
    End If
        If Xcor0 + Xcor1 > Pic1.Width Then Xcor1 = Pic1.Width - Xcor0
        If Ycor0 + Ycor1 > Pic1.Height Then Ycor1 = Pic1.Height - Ycor0
        Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
        SetCoordinates


End Select
End Sub

Private Sub mnuSelection1_Click()
ToolXY.Show
End Sub

Private Sub mnuSelLang_Click(Index As Integer)
On Error Resume Next
     Select Case Index
        Case 0
            If mnuSelLang(Index).Checked = False Then
            Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Lang", "EN")
            Main
            LoadLang1
            End If
        Case 1
             If mnuSelLang(Index).Checked = False Then
            Call savestring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "Lang", "SC")
            Main
            LoadLang1
            End If
    End Select
    Unload FStuLogo
    Unload FWel
 If mnuFiles(1).Enabled = True Then SetPicInfo '刷新文件信息
 SetCoordinates '刷新筐选字
End Sub

Private Sub mnuSoft_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 3
Case 1
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 1
Case 2
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 10
Case 3
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 9
Case 4
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 2
End Select
End Sub

Private Sub mnuHard_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 4
Case 1
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 5
Case 2
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 0
Case 3
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 8
End Select
End Sub

Private Sub mnuSpFil_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Brown Xcor0, Ycor0, Xcor1, Ycor1, 128
Case 1
Brown Xcor0, Ycor0, Xcor1, Ycor1, 256
Case 2
Liquid Xcor0, Ycor0, Xcor1, Ycor1
Case 3
Yellow Xcor0, Ycor0, Xcor1, Ycor1
Case 4
Charcoal Xcor0, Ycor0, Xcor1, Ycor1
Case 5
DarkMoon Xcor0, Ycor0, Xcor1, Ycor1
Case 6
TotalEclipse Xcor0, Ycor0, Xcor1, Ycor1
Case 7
PurpleRain Xcor0, Ycor0, Xcor1, Ycor1
Case 8
Spooky Xcor0, Ycor0, Xcor1, Ycor1
Case 9
UnReal Xcor0, Ycor0, Xcor1, Ycor1
Case 10
Flame Xcor0, Ycor0, Xcor1, Ycor1
Case 11
Aquarel Xcor0, Ycor0, Xcor1, Ycor1
Case 12
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 6
Case 13
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 7
Case 14
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 11
End Select
End Sub

Private Sub mnuSup1_Click()
Fsup.Show 1
End Sub

Private Sub mnuSwap_Click(Index As Integer)
SaveRedo
SwapComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnuThankto_Click()

If Dir(App.Path + "\Studio.dll") = "" Then
MsgBox "Cannot find the Logo File!", vbCritical, "Error"
Else
Load FStuLogo
With FStuLogo
.Timer1.Enabled = False
.Timer2.Enabled = False
.Command1.Visible = True
.Image1.Picture = LoadPicture(App.Path + "\StudioF.dll")
.Label1.Visible = True
.Label2.Visible = True
.Label5.Visible = True
.Show 1
End With
End If
End Sub

Private Sub mnuTxt_Click(Index As Integer)
Select Case Index
Case 0 'add text
FText.Show 1
End Select
End Sub

Private Sub mnuUpdate_Click()
With FAddPT
.Caption = "Check Update"
.Frame1(1).Visible = True
.Show 1
End With
End Sub

Private Sub munReg1_Click()
GY.Show 1
End Sub

Private Sub munWeb1_Click()
ShellExecute Me.hWnd, "Open", lgT(8), "", App.Path, 1
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Shape1.Visible = True
Xcor0 = X: Ycor0 = Y
Xcor1 = 0: Ycor1 = 0
XXX1 = Xcor0: YYY1 = Ycor0
XXX2 = Xcor1: YYY2 = Ycor1
Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
SetCoordinates
FMain.Toolbar1.Buttons(3).Enabled = True
FMain.mnuSel(1).Enabled = True

Pic1.MousePointer = 2
ElseIf Button = 2 Then
SelectAll
End If
End Sub


Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
            XXX2 = X - XXX1: YYY2 = Y - YYY1
            Xcor0 = XXX1: Ycor0 = YYY1
            Xcor1 = XXX2: Ycor1 = YYY2
    If XXX2 < 0 Then
    Xcor0 = XXX1 + XXX2
            If Xcor0 < 0 Then Xcor0 = 0
    Xcor1 = XXX1 - Xcor0
    End If
    If YYY2 < 0 Then
    Ycor0 = YYY1 + YYY2
            If Ycor0 < 0 Then Ycor0 = 0
    Ycor1 = YYY1 - Ycor0
    End If
        If Xcor0 + Xcor1 > Pic1.Width Then Xcor1 = Pic1.Width - Xcor0
        If Ycor0 + Ycor1 > Pic1.Height Then Ycor1 = Pic1.Height - Ycor0
        Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
        SetCoordinates
End If
Pic1.MousePointer = 2
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Xcor1 - Xcor0 = 0 And Ycor1 - Ycor0 = 0 Then
SelectAll
End If
End Sub

Private Sub PicX_Resize()
VS1.Left = PicX.Width - 15
HS1.Top = PicX.Height - 16
HS1.Width = PicX.Width - 15
VS1.Height = PicX.Height - 15
Dummy.Top = PicX.Height - 15
Dummy.Left = PicX.Width - 15
End Sub



Private Sub Run1_Click()
mnuFiles_Click (0)
End Sub

Private Sub Run2_Click()
mnuFiles_Click (0)
End Sub



Private Sub SelectTool_Click(Index As Integer)
Select Case Index
Case 1
mnuFilter_Click 0
Case 2
mnuFilter_Click 3
Case 3
ToolF(0).Visible = True
Case 4
mnuFilter_Click 5
Case 5
mnuFilter_Click 7
Case 6
mnuFilter_Click 10
Case 7
mnuFilter_Click 12
Case 8
PopupMenu Me.mnuFilter(22)

Case 9
mnuFilter_Click 14
Case 10
mnuFilter_Click 15
Case 11
mnuFilter_Click 16
Case 12
mnuFilter_Click 17
Case 13
mnuFilter_Click 19
Case 14
ToolF(1).Visible = True

' Start for SF
Case 15
mnuSpFil_Click 0
Case 16
mnuSpFil_Click 1
Case 17
mnuSpFil_Click 2
Case 18
mnuSpFil_Click 3
Case 19
mnuSpFil_Click 4
Case 20
mnuSpFil_Click 5
Case 21
mnuSpFil_Click 6
Case 22
mnuSpFil_Click 7
Case 23
mnuSpFil_Click 8
Case 24
mnuSpFil_Click 9
Case 25
mnuSpFil_Click 10
Case 26
mnuSpFil_Click 11
Case 27
mnuSpFil_Click 12
Case 28
mnuSpFil_Click 13
Case 29
mnuSpFil_Click 14
' Effects

Case 30
mnuEff_Click 0

Case 31
mnuEff_Click 1
Case 32
mnuEff_Click 2
Case 33
mnuEff_Click 3
Case 34
mnuEff_Click 5
Case 35
mnuEff_Click 6
Case 36
mnuEff_Click 7
Case 37
mnuEff_Click 8
Case 38
mnuEff_Click 9
Case 39
mnuEff_Click 10
Case 40
mnuEff_Click 11
Case 41
mnuEff_Click 12
Case 42
mnuEff_Click 13
Case 43
mnuEff_Click 14
Case 44
mnuEff_Click 15
Case 45
mnuEff_Click 16
Case 46
mnuEff_Click 17
Case 47
mnuEff_Click 18
Case 48
PopupMenu Me.mnuEff(20)

Case 49
mnuPic_Click 0
Case 50
mnuPic_Click 1
Case 51
mnuPic_Click 3
Case 52
mnuPic_Click 4
Case 53
mnuPic_Click 5
Case 54
mnuPic_Click 6

'Def
Case 55
mnuDef_Click 0
Case 56
mnuDef_Click 2
Case 57
mnuDef_Click 3
Case 58
mnuDef_Click 5
Case 59
mnuDef_Click 6
Case 60
mnuDef_Click 7
Case 61
mnuDef_Click 8
Case 62
mnuDef_Click 10
Case 63  ' mix
mnuMix_Click 0
Case 64
mnuMix_Click 1
Case 65
mnuMix_Click 2
Case 66
mnuMix_Click 3
Case 67
mnuMix_Click 4
Case 68
mnuMix_Click 5
Case 69
mnuMix_Click 6
Case 70
mnuMix_Click 8
Case 71
mnuMix_Click 9
' Color
Case 72
mnuCol_Click 0
Case 73
PopupMenu Me.mnuCol(2)
Case 74
PopupMenu Me.mnuCol(4)
Case 79
PopupMenu Me.mnuCol(11)


Case 75
mnuCol_Click 6
Case 76
mnuCol_Click 7
Case 77
mnuCol_Click 9
Case 78
mnuCol_Click 13
End Select
End Sub

Private Sub SelectTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If SelectTool(Index).BackStyle = 0 Then
    For sti = 0 To 100
    SelectTool(sti).BackStyle = 0
    Next
End If
SelectTool(Index).BackStyle = 1
SelectTool(Index).BackColor = &HDCB8DA
End Sub

Private Sub TabBg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmove.Visible = False
Bover.Visible = False
End Sub

Private Sub Timer1_Timer()
If Shape1.Visible = False Then Exit Sub
Tim = (Tim + 1) And 15
Shape1.BorderColor = Scol(Tim)
End Sub


Private Sub Timer2_Timer()
Timer2.Enabled = False
SBar1.Panels(1).Text = lgT(326)
Unload FLoading
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
If newVer = "1" Then
LabNote1.Visible = True
PicNote1.Visible = True
End If
End Sub


Private Sub Timer4_Timer()
HideTip
End Sub

Private Sub Tol1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    For sti = 0 To 100
    SelectTool(sti).BackStyle = 0
    Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "keyUndo"
Redo
Case "keySelectAll"
SelectAll
End Select
End Sub




Private Sub ToolF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For oe = 0 To 20
ImgT2(oe).Borderstyle = 0
Next
End Sub

Private Sub ToolFunc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    For sti = 0 To 100
    SelectTool(sti).BackStyle = 0
    Next
End Sub

Private Sub VS1_GotFocus()
Dummy.SetFocus
End Sub

Private Sub VS1_Scroll()
Pic1.Top = VS1.Value
Pic1.SetFocus
End Sub


Sub ToolShow()
If mnuFilters.Enabled = False Then
Call mnuFiles_Click(0)
    If mnuFilters.Enabled = True Then GoTo ShowOK
Else
GoTo ShowOK
End If
Exit Sub
ShowOK:
On Error Resume Next
For tolBtm = 0 To 8
Command2(tolBtm).Visible = False
Image23(tolBtm).Visible = False
Next
CmdBack.Visible = True
ToolUseImg.Visible = True
ToolName.Visible = True
ToolName2.Visible = True


End Sub

Sub ToolHide()
On Error Resume Next
For tolBtm = 0 To 8
Command2(tolBtm).Visible = True
Image23(tolBtm).Visible = True
Next
CmdBack.Visible = False
ToolUseImg.Visible = False
ToolName.Visible = False
ToolName2.Visible = False


'Add more!!              Append Design Function Display
For TFh1 = 0 To 6
'隐藏面板
ToolFunc(TFh1).Visible = False
Next
End Sub


Sub HideTip()
Image5.Visible = False
Label9.Visible = False
Label11.Visible = False
Shape2.Visible = False
Line3.Visible = False
Line4.Visible = False
End Sub
