VERSION 5.00
Begin VB.UserControl Xp_Frame 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   ControlContainer=   -1  'True
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   62
   ToolboxBitmap   =   "FrameXp.ctx":0000
   Begin VB.Image img 
      Height          =   300
      Left            =   0
      Picture         =   "FrameXp.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " Frame1"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   24
      Top             =   0
      Width           =   570
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   0
      Left            =   9000
      TabIndex        =   23
      Top             =   240
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   22
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   21
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   9000
      TabIndex        =   20
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   9000
      TabIndex        =   19
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   18
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   6
      Left            =   9000
      TabIndex        =   17
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   16
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   8
      Left            =   9000
      TabIndex        =   15
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   9
      Left            =   9000
      TabIndex        =   14
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   10
      Left            =   9000
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   11
      Left            =   9000
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   12
      Left            =   9000
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   13
      Left            =   9000
      TabIndex        =   10
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   14
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   15
      Left            =   9000
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   16
      Left            =   9000
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   17
      Left            =   9000
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   18
      Left            =   9000
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   19
      Left            =   9000
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   20
      Left            =   9000
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   21
      Left            =   9000
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   22
      Left            =   9000
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label dot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   23
      Left            =   9000
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "Xp_Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*   *  This code is originaly writen by Aleks Kimhi - Aki.                                               *
'  *    As you might have read that this ActiveX control is free,                 O                       *
'*   *  and therefore can be distributed for free as long as you                                              *
'  *    do not sell it for profit and there's still my name on it.                               ____|            *
'*   *  I've spend some time on this project so please                                                            *
'  *    don't just take it, use it and say you wrote it.                                                               *
'*   *  Thank you for your co-operation. Any comments,                           P                          *
'  *    good or bad, would be greatly appreciated.                                                               *
'*   *  E -mail: aniram@ zahav.net.il                                                                               *
'  *    Tel-Aviv, Israel    A & M © Copyright 2002                                                           *
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
' *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Dim mBackColor As OLE_COLOR                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                'Aki
Dim mForeColor As OLE_COLOR
Dim mFont As Font
Const defBackColor = vbButtonFace
Const defForeColor = &HFF0000
'In this project I tried to use only one picture to make XpFrame but it was not that.
'The corners were like Microsoft Frame and I wanted XpFrame,so that was not the solution.
'Other solution was to use labels with caption " . ", only for corners.
'I had a problem when you change font to other size, but I fix that to, so everything
'is working perfect and you have now the final version of XpFrame->>by Aki.

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    BackColor = defBackColor
    Caption = Ambient.DisplayName
    Set Font = UserControl.Ambient.Font
    ForeColor = defForeColor
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 3
    UserControl.Cls
    
    Dim x, Y, W, H, lblH As Integer
        x = UserControl.ScaleWidth - 3
        Y = UserControl.ScaleHeight - 3
        W = UserControl.ScaleWidth - 6
        H = UserControl.ScaleHeight - 6 - (lbl.Height \ 2)
        lblH = lbl.Height \ 2
        
            If lbl.Caption = "" Then
                lbl.Visible = False
                    Else
                        lbl.Visible = True
            End If
                lbl.Top = 0
                    lbl.Left = 12
    
    UserControl.PaintPicture img.Picture, 3 + 3, lblH, W - 6, 1, 3, 0, 16, 1 'painting top line
    UserControl.PaintPicture img.Picture, x + 2, lblH + 6, 1, H - 6, 21, 3, 1, 14 'painting right line
    UserControl.PaintPicture img.Picture, 3 + 3, Y + 2, W - 6, 1, 3, 19, 16, 1 'painting bottom line
    UserControl.PaintPicture img.Picture, 0, lblH + 6, 1, H - 6, 0, 3, 1, 14 'painting left line
    
    'starting to paint corners using labels with caption " . " --> ONLY corners
    'to look perfect you must know that every corner needs 6 dots, so here we go...
        
    'painting left top
    dot(0).Top = lblH - 5: dot(0).Left = UserControl.ScaleWidth - UserControl.ScaleWidth
        dot(1).Top = lblH - 6: dot(1).Left = UserControl.ScaleWidth - UserControl.ScaleWidth
            dot(2).Top = lblH - 7: dot(2).Left = UserControl.ScaleWidth - UserControl.ScaleWidth + 1
                dot(3).Top = lblH - 8: dot(3).Left = UserControl.ScaleWidth - UserControl.ScaleWidth + 2
                    dot(4).Top = lblH - 9: dot(4).Left = UserControl.ScaleWidth - UserControl.ScaleWidth + 3
                        dot(5).Top = lblH - 9: dot(5).Left = UserControl.ScaleWidth - UserControl.ScaleWidth + 4
                            'painting right top
                            dot(6).Top = lblH - 5: dot(6).Left = UserControl.ScaleWidth - 3
                                dot(7).Top = lblH - 6: dot(7).Left = UserControl.ScaleWidth - 3
                                    dot(8).Top = lblH - 7: dot(8).Left = UserControl.ScaleWidth - 4
                                        dot(9).Top = lblH - 8: dot(9).Left = UserControl.ScaleWidth - 5
                                            dot(10).Top = lblH - 9: dot(10).Left = UserControl.ScaleWidth - 6
                                                dot(11).Top = lblH - 9: dot(11).Left = UserControl.ScaleWidth - 7
                                                    'painting right down
                                                    dot(12).Top = UserControl.ScaleHeight - 16: dot(12).Left = UserControl.ScaleWidth - 3
                                                dot(13).Top = UserControl.ScaleHeight - 15: dot(13).Left = UserControl.ScaleWidth - 3
                                            dot(14).Top = UserControl.ScaleHeight - 14: dot(14).Left = UserControl.ScaleWidth - 4
                                        dot(15).Top = UserControl.ScaleHeight - 13: dot(15).Left = UserControl.ScaleWidth - 5
                                    dot(16).Top = UserControl.ScaleHeight - 12: dot(16).Left = UserControl.ScaleWidth - 6
                                dot(17).Top = UserControl.ScaleHeight - 12: dot(17).Left = UserControl.ScaleWidth - 7
                            'painting left down
                            dot(18).Top = UserControl.ScaleHeight - 16: dot(18).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth)
                        dot(19).Top = UserControl.ScaleHeight - 15: dot(19).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth)
                    dot(20).Top = UserControl.ScaleHeight - 14: dot(20).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth) + 1
                dot(21).Top = UserControl.ScaleHeight - 13: dot(21).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth) + 2
            dot(22).Top = UserControl.ScaleHeight - 12: dot(22).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth) + 3
        dot(23).Top = UserControl.ScaleHeight - 12: dot(23).Left = (UserControl.ScaleWidth - UserControl.ScaleWidth) + 4
End Sub

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    BackColor = PropBag.ReadProperty("BackColor", defBackColor)
    Caption = PropBag.ReadProperty("Caption", "Frame1")
    ForeColor = PropBag.ReadProperty("ForeColor", defForeColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "Frame")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", mFont, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("BackColor", mBackColor, defBackColor)
    Call PropBag.WriteProperty("ForeColor", mForeColor, defForeColor)
End Sub

Public Property Get Font() As Font
    Set Font = mFont
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set mFont = NewFont
    Set lbl.Font = mFont
    UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    mForeColor = NewForeColor
    lbl.ForeColor = mForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    mBackColor = NewBackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = mBackColor
    lbl.BackColor = mBackColor
    UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lbl.Caption() = NewCaption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    UserControl.Enabled() = NewEnabled
        If Enabled = True Then
            lbl.ForeColor = &HFF0000
                Else
                    lbl.ForeColor = &H80000011
        End If
    PropertyChanged "Enabled"
End Property
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           'Aki
