VERSION 5.00
Begin VB.UserControl dEditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   765
      Top             =   615
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   645
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox pLines 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E6E6E6&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   0
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   645
      Begin VB.Line lbSpacer 
         BorderColor     =   &H8000000C&
         X1              =   42
         X2              =   42
         Y1              =   0
         Y2              =   35
      End
   End
End
Attribute VB_Name = "dEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_GETLINE As Long = &HC4
Private Const EM_LINEINDEX = &HBB
Private Const EM_UNDO = &HC7
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
'Event Declarations:
Event Change()

Public Sub SetActive()
On Error Resume Next
    Call txtCode.SetFocus
End Sub

Public Sub Undo()
    SendMessage txtCode.hwnd, EM_UNDO, ByVal 0&, ByVal 0&
End Sub

Public Sub SelectAll()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
    Call txtCode.SetFocus
End Sub

Public Sub Cut()
On Error Resume Next
   SendMessage txtCode.hwnd, WM_CUT, 0, ByVal 0&
    Call txtCode.SetFocus
End Sub

Public Sub Copy()
On Error Resume Next
    SendMessage txtCode.hwnd, WM_COPY, 0, ByVal 0&
    Call txtCode.SetFocus
End Sub

Public Sub Paste()
On Error Resume Next
    txtCode.SelText = Replace(Clipboard.GetText(vbCFText), vbTab, String$(8, " "))
    Call txtCode.SetFocus
End Sub

Private Sub DrawLines()
Dim Counter As Long
Dim sLine As String

    'This sub draws the line numbers
    With pLines
        'Clear DC
        Call .Cls
        Set .Font = txtCode.Font
        For Counter = (GetVisableLine + 1) To GetLineCount
            'Set normal text color
            .ForeColor = vbBlack
            .CurrentX = (.Width - 10) - .TextWidth(Str$(Counter))
            If (Counter = LineIndex) Then
                'Set line heighlight color
                .ForeColor = &H808080
            End If
            'Print lines
            pLines.Print Counter
        Next Counter
    End With
    
End Sub

Private Function LineIndex() As Long
    LineIndex = SendMessage(txtCode.hwnd, EM_LINEFROMCHAR, (txtCode.SelStart + txtCode.SelLength), 0) + 1
End Function

Private Function GetLineCount() As Long
    GetLineCount = SendMessage(txtCode.hwnd, EM_GETLINECOUNT, 0, 0)
End Function

Private Function GetVisableLine() As Long
    GetVisableLine = SendMessage(txtCode.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Function

Private Sub pLines_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim ln As Integer
On Error Resume Next
    If Button = vbLeftButton Then
        'Calc line
        ln = (y \ pLines.TextHeight("Az"))
        txtCode.SelStart = SendMessage(txtCode.hwnd, EM_LINEINDEX, ln, 0)
        Call txtCode.SetFocus
    End If
End Sub

Private Sub Tmr_Timer()
    Call DrawLines
End Sub

Private Function GetNoneSpaceIdx(s0 As String) As Integer
Dim I As Integer
Dim Idx As Integer
    
    For I = 1 To Len(s0)
        If Mid(s0, I, 1) <> " " Then
            Idx = I
            Exit For
        End If
    Next I
    
    GetNoneSpaceIdx = Idx
    
End Function

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Dim Lines() As String
Dim Idx As Integer
On Error Resume Next
    'Select all text
    Dim PrevLn As Integer
    
    PrevLn = SendMessage(txtCode.hwnd, EM_LINEFROMCHAR, -1, 0)
    
    'Auto tab
    If KeyAscii = 13 Then
        Lines = Split(txtCode.Text, vbCrLf)
        If PrevLn = 0 And Len(txtCode.Text) = 0 Then Exit Sub
        'Get size of spaces
        Idx = GetNoneSpaceIdx(Lines(PrevLn))
        
        If (Idx > 1) Then
            txtCode.SelText = vbCrLf & String$(Idx, " ")
            txtCode.SelStart = txtCode.SelStart - 1
            KeyAscii = 0
        End If
        Erase Lines
    End If
    
    'Auto close "
    If KeyAscii = 34 Then
        txtCode.SelText = """"
        txtCode.SelStart = txtCode.SelStart - 1
    End If
    
    'Auto close {
    If KeyAscii = 123 Then
        txtCode.SelText = "}"
        txtCode.SelStart = txtCode.SelStart - 1
    End If
    
    If (KeyAscii = 1) Then
        txtCode.SelStart = 0
        txtCode.SelLength = Len(txtCode.Text)
        txtCode.SetFocus
        KeyAscii = 0
    End If
    If (KeyAscii = 9) Then
        txtCode.SelText = Space(4)
        KeyAscii = 0
    End If
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    pLines.Height = (UserControl.ScaleHeight)
    lbSpacer.Y2 = pLines.ScaleHeight
    'Resize editor
    txtCode.Height = (UserControl.ScaleHeight - txtCode.Top)
    txtCode.Width = (UserControl.ScaleWidth - txtCode.Left)
End Sub

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtCode.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtCode.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtCode.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtCode.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtCode.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtCode.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtCode.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtCode.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtCode.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtCode.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtCode.Locked = PropBag.ReadProperty("Locked", False)
    txtCode.Text = PropBag.ReadProperty("Text", "Text1")
    txtCode.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set txtCode.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtCode.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtCode.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtCode.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtCode.SelText = PropBag.ReadProperty("SelText", "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

Private Sub UserControl_Show()
    Tmr.Enabled = (UserControl.Ambient.UserMode)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Locked", txtCode.Locked, False)
    Call PropBag.WriteProperty("Text", txtCode.Text, "Text1")
    Call PropBag.WriteProperty("BackColor", txtCode.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", txtCode.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", txtCode.ForeColor, &H80000008)
    Call PropBag.WriteProperty("SelLength", txtCode.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtCode.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtCode.SelText, "")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtCode.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtCode.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtCode.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtCode.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtCode.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtCode.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Private Sub txtCode_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

