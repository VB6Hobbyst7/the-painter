VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "ScreenGrabForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
  Dim z As POINTAPI
  Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long



Dim xStart As Single, yStart As Single, bMouseDown As Boolean
Dim xs, ys, xs2, ys2

Private Sub Form_Load()
If FTimerQs.Check1.Value = 1 Then
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE _
Or SWP_NOSIZE
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIForm1.Visible = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'start of mouse down coords xStart:yStart
    xStart = X: yStart = Y
    
    bMouseDown = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    If bMouseDown = True Then

        movelines X, Y
        
        xs = X: ys = Y
        Dim xbig, ybig
        'place label info at bottom right;get biggest coords
        xbig = max(xStart, xs)
        ybig = max(yStart, ys)
        
        Label1.Visible = False
        Label1.Left = xbig + 4: Label1.Top = ybig + 4
        
        Label1.Caption = "X=" & Format$(X, "0000") & vbCrLf & "Y=" & Format$(Y, "0000") & vbCrLf & "Width=" & Format$(Abs(X - xStart), "0000") _
           & vbCrLf & "Height=" & Format$(Abs(Y - yStart), "0000")
        
        Label1.Visible = True

    End If

    'Form1.Caption = "X= " & Format$(X, "0000") & ": Y= " & Format$(Y, "0000")
    Form1.Caption = Format$(X, "0000") & ":" & Format$(Y, "0000") & ":" & Format$(Abs(X - xStart), "0000") _
       & ":" & Format$(Abs(Y - yStart), "0000")

End Sub

Sub movelines(X As Single, Y As Single)

    If Not (xs = 0 And ys = 0) Then

        'delete previous
        '''-Form1.Line (xStart, yStart)-(xs - 1, ys - 1), , B
        Form1.Line (xStart, yStart)-(xs, ys), , B

    End If

    'draw selection square in invert drawmode
    '''-Form1.Line (xStart, yStart)-(x - 1, y - 1), , B
    Form1.Line (xStart, yStart)-(X, Y), , B

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errMouseUp
    ''Shape1.Visible = False
    Label1.Visible = False

    bMouseDown = False
    ''Form1.Line (xStart, yStart)-(xs, ys), , B
    '''Form1.Line (xStart, yStart)-(xs, ys), , B
    'delete previous
    Form1.Line (xStart, yStart)-(xs, ys), , B

    Form1.Line (xs, 0)-(xs, ys - 10) '(10 + Shape1.Width))
   
    Dim xwidth, yheight
    Dim startx, starty
    Dim endx, endy
    xwidth = Abs(xStart - xs)
    yheight = Abs(yStart - ys)
    If debugme = True Then MsgBox "xStart = " & xStart & "yStart= " & yStart
    If debugme = True Then MsgBox xwidth & ":" & yheight
    'if mouse start x/y positions = new x/y positions
    If xStart = X And yStart = Y Then
        If debugme = True Then MsgBox "xStart =x and yStart=y"
        xs = 0: ys = 0
        Unload Me
        'stops rest of code executing
        Exit Sub
    End If
    'get new form to blit to
    If xwidth <= 0 Or yheight <= 0 Then
        MsgBox "No Pic width or height"
        Exit Sub
    End If
    'create new child forms of MDI
    Dim frmChild As New frmChild
    frmChild.Show

    If MDIForm1.ActiveForm Is Nothing Then
    'somehow we have no child form
        MsgBox "need form to blit to"
        Exit Sub
    End If

    frmChild.Picture1.Visible = False

    With MDIForm1.ActiveForm.Picture1

        .BackColor = &HFF00FF
        .Cls
        ''
        '.Width = xwidth + 150
        ''.Width = Screen.TwipsPerPixelX * (xwidth + 8)
        .Width = xwidth + 1

        If debugme = True Then MsgBox .Width

        '''.Width = xwidth 'Shape1.Width
        ''.Height = yheight + 150 'Shape1.Height
        ''.Height = Screen.TwipsPerPixelY * (yheight + 26)
        .Height = yheight + 1

        If debugme = True Then MsgBox .Height

        'systemmetrics 26= caption and menubar;8= 3d borders of form
        MDIForm1.ActiveForm.Width = Screen.TwipsPerPixelX * (xwidth + 40)
        MDIForm1.ActiveForm.Height = Screen.TwipsPerPixelY * (yheight + 50)
        ''    '     '
        ''get the correct coords;swap if need be
        'draw from top left corner down to right
        If xStart <= xs And yStart <= ys Then

            startx = xStart: starty = yStart

        End If

        ''draw from bottom right to top left
        If xStart > xs And yStart > ys Then
            startx = xs: starty = ys
        End If

        ''from bottom left to top right
        If xStart < xs And yStart > ys Then
            startx = xStart
            starty = yStart - yheight
        End If

        ''from bottom right to top left
        If xStart > xs And yStart < ys Then
            startx = xStart - xwidth
            starty = yStart
        End If
        '''If xStart > xs Then
        'copy from grab screen form (form1) to

        If xwidth > 0 And yheight > 0 Then
            MDIForm1.ActiveForm.Picture1.PaintPicture Form1.Picture, 0, 0, , , startx, starty, xwidth + 1, yheight + 1
        End If

        .Visible = True

    End With

    xs = 0: ys = 0
    'unload me?
    'convert picture
    MDIForm1.ActiveForm.Picture1.Picture = MDIForm1.ActiveForm.Picture1.Image
    frmChild.Picture1.Visible = True
    MDIForm1.WindowState = 0
    Unload Me
    Exit Sub

errMouseUp:
    xs = 0: ys = 0
    MsgBox Err.Description & ": Error number " & Err.Number

End Sub


Sub GrabNow1()

End Sub
