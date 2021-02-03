Attribute VB_Name = "MainMod1"
Public LangA
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '执行文件的声明
   Type POINTAPI
    X As Long
    Y As Long
End Type

Sub Main()
On Error Resume Next
        LangA = "lgc"


If Dir(App.Path + "\Studio.exe") = "" Then

    MsgBox "安装不完整，请重新运行安装程序", vbCritical, "Serious"
    End
Else

    If Command = "" Then
        MDIForm1.Show
    ElseIf Command = "cs" Then
        frmColor.Show
    ElseIf Command = "sg" Then
        MDIForm1.Show
    End If

End If

End Sub
