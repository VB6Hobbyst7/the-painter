VERSION 5.00
Begin VB.Form ToolZoom 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   3045
   ClientLeft      =   75
   ClientTop       =   6105
   ClientWidth     =   4065
   LinkTopic       =   "Form5"
   ScaleHeight     =   3045
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrZoom 
      Interval        =   50
      Left            =   1830
      Top             =   720
   End
   Begin VB.HScrollBar hsbZoom 
      Height          =   270
      LargeChange     =   10
      Left            =   480
      Max             =   1000
      Min             =   25
      TabIndex        =   5
      Top             =   350
      Value           =   25
      Width           =   2760
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2380
      Left            =   50
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   4
      Top             =   620
      Width           =   3970
      Begin VB.Line Line1 
         X1              =   124
         X2              =   140
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Line Line2 
         X1              =   132
         X2              =   132
         Y1              =   72
         Y2              =   88
      End
   End
   Begin VB.CheckBox chkGrid 
      Height          =   270
      Left            =   120
      Picture         =   "ToolZoom.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   350
      Width           =   240
   End
   Begin VB.CheckBox chkOnTop 
      DownPicture     =   "ToolZoom.frx":00F6
      Enabled         =   0   'False
      Height          =   270
      Left            =   3240
      Picture         =   "ToolZoom.frx":01EC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtZoom 
      Height          =   285
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "1000%"
      Top             =   350
      Width           =   570
   End
   Begin VB.Image IList2 
      Height          =   210
      Left            =   2760
      Picture         =   "ToolZoom.frx":02E2
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IList1 
      Height          =   210
      Left            =   2520
      Picture         =   "ToolZoom.frx":05C4
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Ialpha2 
      Height          =   135
      Left            =   2040
      Picture         =   "ToolZoom.frx":08A6
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Ialpha1 
      Height          =   135
      Left            =   1800
      Picture         =   "ToolZoom.frx":09E4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   3720
      MouseIcon       =   "ToolZoom.frx":0B22
      MousePointer    =   99  'Custom
      Picture         =   "ToolZoom.frx":0C74
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   120
      Picture         =   "ToolZoom.frx":0F56
      Top             =   100
      Width           =   135
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
      Height          =   255
      Left            =   120
      MouseIcon       =   "ToolZoom.frx":1094
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Width           =   1440
   End
   Begin VB.Image Image10 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   4020
      Picture         =   "ToolZoom.frx":11E6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image11 
      Height          =   45
      Left            =   0
      Picture         =   "ToolZoom.frx":1534
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   4080
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   0
      MousePointer    =   15  'Size All
      Picture         =   "ToolZoom.frx":1EE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
   Begin VB.Image Image8 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   0
      Picture         =   "ToolZoom.frx":664C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "ToolZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'User Defined Types
Private Type PointAPI   'API point structure.
    X   As Long
    Y   As Long
End Type

Private Type SizeRect   'Size structure (uses Width, Height instead of bounds)
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type RectAPI    'Rect structure (uses Right, Bottom bounds instead of Width, Height)
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Windows API Blt (BitBlt, PatBlt, StretchBlt) ROP constants.
Private Const SRCCOPY           As Long = &HCC0020
Private Const PATCOPY           As Long = &HF00021

'SetWindowPos Flags.
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FLAGS         As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

'Module level variables.
Private mfScale As Single   'Scale of Zoom percentage (6 = 600%) (6 x Size = 600% increase)
Private mlOldX  As Long     'Holds Last X-coord of mouse
Private mlOldY  As Long     'Holds Last Y-coord of mouse

'Declare the Windows API functions that are to be used.
'Alphabetical order to ease lookup later.
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Dim WinPo As Boolean
Private Function CreateCheckeredBrush(ByVal hDC As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long

Dim X           As Long
Dim Y           As Long
Dim lRet        As Long
Dim hBitmapDC   As Long
Dim hBitmap     As Long
Dim hOldBitmap  As Long
    
    'Convert System Colors if needed
    If lColor1 < 0 Then
        lColor1 = GetSysColor(lColor1 And &HFF&)
    End If
    If lColor2 < 0 Then
        lColor2 = GetSysColor(lColor2 And &HFF&)
    End If
    
    'Create a new DC and Bitmap to draw the Brush
    hBitmapDC = CreateCompatibleDC(hDC)
    hBitmap = CreateCompatibleBitmap(hDC, 8, 8)
    'Select the Bitmap into the DC for drawing
    hOldBitmap = SelectObject(hBitmapDC, hBitmap)
    
    'Draw the Brush's Bitmap (Checkerboard)
    For Y = 0 To 6 Step 2
        For X = 0 To 6 Step 2
            lRet = SetPixelV(hBitmapDC, X, Y, lColor1)
            lRet = SetPixelV(hBitmapDC, X + 1, Y, lColor2)
            lRet = SetPixelV(hBitmapDC, X, Y + 1, lColor2)
            lRet = SetPixelV(hBitmapDC, X + 1, Y + 1, lColor1)
        Next X
    Next Y
    
    'Get the bitmap back out of the DC
    hBitmap = SelectObject(hBitmapDC, hOldBitmap)
    
    'Create the Brush from the bitmap
    CreateCheckeredBrush = CreatePatternBrush(hBitmap)
    
    'Delete the DC and Bitmap to free memory
    lRet = DeleteDC(hBitmapDC)
    lRet = DeleteObject(hBitmap)

End Function


Private Sub DoZoom(ptMouse As PointAPI)

Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect

    'Get the Desktop DC
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    
    'Setup the Destination size for StretchBlt.
    With sizDest
        .Left = 0
        .Top = 0
        .Width = picZoom.ScaleWidth
        .Height = picZoom.ScaleHeight
    End With
    
    'Setup the Source size for StretchBlt.
    With sizSrce
        .Left = ptMouse.X - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.Y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        'Adjust Source and Destination sizes if they don't match.
        'sizSrce.Size * mfScale must= sizDest.Size for acurate scaling.
        'Destination must always be as large or larger than picZoom.
        'Adjust the Width, if needed.
        lTemp = Int(.Width * mfScale)  '(Source.Width * mfScale must= sizDest.Width)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        'Adjust the Height, if needed.
        lTemp = Int(.Height * mfScale) '(sizSrce.Height * mfScale must= sizDest.Height)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    
    'Clear the current contents.
    picZoom.Cls
    
    'Stretch the Desktop (source) into picZoom (dest)
    lRet = StretchBlt(picZoom.hDC, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    
    'Release the Desktop DC
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    
    'Redraw the grid, if needed
    If chkGrid.Value = vbChecked Then
        Call DrawGrid
    End If
    
    picZoom.Refresh
    
End Sub

Private Sub DrawGrid()

Dim iWidth      As Integer
Dim iHeight     As Integer
Dim lRet        As Long
Dim hBrush      As Long
Dim hOldBrush   As Long
Dim fX          As Single
Dim fY          As Single

    If mfScale >= 3 Then
    
        'Create a Checkered Brush (Dark and Light Grey)...
        hBrush = CreateCheckeredBrush(picZoom.hDC, &H808080, &HC0C0C0)
        '...and Select it into the PictureBox
        hOldBrush = SelectObject(picZoom.hDC, hBrush)
        
        iWidth = picZoom.ScaleWidth
        iHeight = picZoom.ScaleHeight
        
        'Draw the gridlines using the checkered pattern brush.
        For fX = 0 To iWidth Step mfScale
            lRet = PatBlt(picZoom.hDC, Int(fX), 0, 1, iHeight, PATCOPY)
        Next
        For fY = 0 To iHeight Step mfScale
            lRet = PatBlt(picZoom.hDC, 0, Int(fY), iWidth, 1, PATCOPY)
        Next
        
        'Put the old Brush back and Delete the new one to free memory
        hBrush = SelectObject(picZoom.hDC, hOldBrush)
        lRet = DeleteObject(hBrush)
    
    End If
    
End Sub
Private Function ValidScale(ByVal fScale As Single) As Single

    'If the user typed an invalid scale,
    'change it to be within Zoom bounds.
    If fScale * 100 > hsbZoom.Max Then
        fScale = hsbZoom.Max / 100
    ElseIf fScale * 100 < hsbZoom.min Then
        fScale = hsbZoom.min / 100
    End If
    
    ValidScale = fScale
    
End Function

Private Sub LoadSettings()

    'Load the saved settings from the init file.
  '  Call RestoreFormSize(Me)
    hsbZoom.Value = GetInitEntry("Settings", "Zoom", CStr(200))
   hsbZoom_Change
    chkGrid.Value = IIf(LCase$(GetInitEntry("Settings", "Grid", "False")) = "true", vbChecked, vbUnchecked)
    chkGrid_Click


End Sub

Private Sub SaveSettings()

Dim lRet As Long

    'Save the current settings to the init file.
  '  Call SaveFormSize(Me)
    lRet = SetInitEntry("Settings", "Zoom", hsbZoom.Value)
    lRet = SetInitEntry("Settings", "Grid", CStr(chkGrid.Value = vbChecked))
    

End Sub

Private Sub Form_Load()
 rtn1 = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
 Me.Left = Screen.Width - Me.Width - 200
 
   Call LoadSettings
WinPo = True

   'Cross
'Dim w1, h1
'w1 = picZoom.Width / 1000
'h1 = picZoom.Height / 1000
'With Line1
'.x1 = w1 / 2 - 11
'.X2 = w1 / 2 + 5
'.y1 = h1 / 2 - 3
'.Y2 = .y1
'End With

'With Line2
'.x1 = w1 / 2 - 3
'.X2 = .x1
'.y1 = h1 / 2 - 11
'.Y2 = h1 / 2 + 5
'
'End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Image4.Picture = IList1.Picture
End Sub

Private Sub Image2_DblClick()
Label3_Click
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
Image4.Picture = IList1.Picture
End Sub

Private Sub Image4_Click()
ActiveTool = 2
PopupMenu FMain.mnuTool, , Image4.Left, Image4.Top + Image4.Height
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = IList2.Picture
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = IList1.Picture
End Sub




Private Sub Label3_Click()
If WinPo = True Then
Me.Height = 330
Image3.Picture = Ialpha2.Picture
WinPo = False
Else
Me.Height = 3045
Image3.Picture = Ialpha1.Picture
WinPo = True
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = True
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = False
End Sub


Private Sub chkGrid_Click()

    'Force the zoom to update.
    mlOldX = -100
    
    'Remove focus from button so there's no focus rect.
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

  Call SaveSettings
    
End Sub


Private Sub hsbZoom_Change()

    'Update the label
    txtZoom.Text = Format$(hsbZoom.Value / 100, "####%")
    
    'Reset mfScale
    mfScale = CSng(hsbZoom.Value) / 100!
    
    'Remove focus from scrollbar so there's no flashing thumb.
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
    'Force the zoom to update
    mlOldX = -100

End Sub


Private Sub hsbZoom_Scroll()

    hsbZoom_Change
    
End Sub




Private Sub tmrZoom_Timer()

Dim lRet    As Long
Dim ptMouse As PointAPI

Static lElapsed As Long

    If Me.WindowState <> vbMinimized Then
        'This code runs 20 times/second*, while the form is not minimized.
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.X <> mlOldX) Or (.Y <> mlOldY) Or (lElapsed >= 250) Then
                'This code runs runs 4 times/second* if no mousemove,
                'or 20 times/second* when mouse is moving.
                Call DoZoom(ptMouse)
                If lElapsed >= 250 Then
                    'This code only runs 4 times/second*.
                    If chkOnTop.Value = vbChecked Then
                        lRet = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
                    End If
                End If
                lElapsed = 0
            End If
            mlOldX = .X
            mlOldY = .Y
        End With
    End If
    
    '* Times/second depends on processor speed. A slower processor may not
    'finish processing one timer event before the next arrives, in which
    'case the new event will be discarded.
    
End Sub


Private Sub txtZoom_GotFocus()

    With txtZoom
        'Remove the "%"
        .Text = CStr(Val(.Text))
        'Select the entire string
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtZoom_KeyPress(KeyAscii As Integer)

    'Allow Numbers and Edit Keys (Backspace) only.
    'Backspace [Asc(8)] won't be affected by this code.
    'Other Edit Keys (Delete, Home, End, PageUp, etc.) fire only
    'KeyDown/KeyUp events and also won't be affected by this code.
    If KeyAscii > 31 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        'Not a number key.
        Beep
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then  '(Asc(13))
        'Force a zoom update, then reselect the textbox. (see _LostFocus)
        picZoom.SetFocus
        DoEvents
        txtZoom.SetFocus
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtZoom_LostFocus()

    'Reset the scale
    mfScale = ValidScale(Val(txtZoom.Text) / 100)
    
    'Update the scrollbar (only fires change event if value changes).
    hsbZoom.Value = mfScale * 100
    
    'Update the textbox in case the scrollbar change event didn't fire.
    txtZoom.Text = Format$(mfScale, "####%")

End Sub



