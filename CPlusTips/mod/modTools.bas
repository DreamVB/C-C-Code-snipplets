Attribute VB_Name = "modTools"
Option Explicit

Public SourceCode As String

Private Const SW_NORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2

Public Sub OpenUrl(Url As String)
Dim RetVal As Long
    RetVal = ShellExecute(GetDesktopWindow(), "open", Url, vbNullString, vbNullString, SW_NORMAL)
End Sub

Public Sub FindListItem(lb As ListBox, TextFind As String)
Dim Ret As Long
On Error Resume Next

    Ret = SendMessage(lb.hwnd, LB_FINDSTRING, -1, ByVal TextFind)
    If Ret <> -1 Then
        lb.ListIndex = Ret
    End If
End Sub
