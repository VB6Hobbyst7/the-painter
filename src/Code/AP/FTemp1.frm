VERSION 5.00
Begin VB.Form FTemp1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FTemp1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Limit1 
      Height          =   270
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox a3 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox a2 
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox a1 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox abc 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "FTemp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '执行文件的声明

