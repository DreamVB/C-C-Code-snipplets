VERSION 5.00
Begin VB.UserControl dmFlatButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   Begin VB.PictureBox PicImg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3165
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "dmFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum TTextAlign
    taLeft = 0
    taCenter = 1
    taRight = 2
End Enum

Enum TPictureAlign
    iaLeft = 0
    iaCenter = 1
    iaRight = 2
End Enum


Enum TDisplay
    TextOnly = 0
    ImageOnly = 1
    TextAndImage = 2
End Enum

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private pt As POINTAPI

Private m_MouseIn As Boolean
Private m_MouseDown As Boolean
Private m_GotFocus As Boolean
Private m_ShowFocus As Boolean

Private m_BackColorHover As OLE_COLOR
Private m_BackColorDefault As OLE_COLOR
Private m_BackgroundMouseDown As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_TextColorDown As OLE_COLOR
Private m_HoverTextColor As OLE_COLOR
Private m_BorderHover As OLE_COLOR
Private m_TxtAlign As TTextAlign
Private m_PictureAlign As TPictureAlign
Private m_Display As TDisplay
Private mMouseButton As MouseButtonConstants

Private m_Text As String
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event DblClick()

Private Sub DrawButton()
Dim xOff As Integer, yOff As Integer
Dim ImgWidth As Integer, ImgHeight As Integer
Dim ImgPosX As Integer, ImgPosY As Integer

Dim rc As RECT

    xOff = 0
    yOff = 0
    
    With UserControl
        .Cls
        
        ImgWidth = PicImg.ScaleWidth
        ImgHeight = PicImg.ScaleHeight
        
        If m_MouseIn Then
            
            If m_MouseDown Then
                .BackColor = MouseDownBackground
            Else
                'Do hover color
                UserControl.BackColor = BackgroundHover
            End If
            
            'Draw border
            UserControl.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), BorderHover, B
            
            If m_MouseDown Then
                .ForeColor = TextColorDown
            Else
                .ForeColor = TextColorHover
            End If
            
        Else
            UserControl.BackColor = BackGround
            .ForeColor = TextColor
        End If
        
        If m_GotFocus And ShowFocusRect Then
            rc.Left = 4
            rc.Right = .ScaleWidth - 4
            rc.Top = 4
            rc.Bottom = .ScaleHeight - 4
            DrawFocusRect .hdc, rc
        End If
        
        'Position image
        Select Case PictureAlign
            Case iaLeft
                ImgPosX = 4
                ImgPosY = (.ScaleHeight - ImgHeight) / 2
            Case iaCenter
                ImgPosX = (.ScaleWidth - ImgWidth) / 2
                ImgPosY = (.ScaleHeight - ImgHeight) / 2
            Case iaRight
                ImgPosX = (.ScaleWidth - ImgWidth - 4)
                ImgPosY = (.ScaleHeight - ImgHeight) / 2
        End Select
        
        If Display = ImageOnly Or Display = TextAndImage Then
            TransparentBlt .hdc, ImgPosX, ImgPosY, ImgWidth, ImgHeight, PicImg.hdc, 0, 0, _
            ImgWidth, ImgHeight, vbMagenta
        End If
        
        If (Display = TextOnly) Or (Display = TextAndImage) Then
            Select Case TextAlign
                Case taLeft
                    .CurrentX = 4
                    .CurrentY = (.ScaleHeight - .TextHeight(Text)) / 2
                Case taCenter
                    .CurrentX = (.ScaleWidth - .TextWidth(Text)) / 2
                    .CurrentY = (.ScaleHeight - .TextHeight(Text)) / 2
                Case taRight
                    .CurrentX = (.ScaleWidth - .TextWidth(Text) - 4)
                    .CurrentY = (.ScaleHeight - .TextHeight(Text)) / 2
            End Select
            
            TextOut .hdc, .CurrentX, .CurrentY, Text, Len(Text)
            
            'UserControl.Print Text
        End If
        
    End With
    
End Sub

Private Sub Timer1_Timer()
Dim iHwnd As Long
    GetCursorPos pt
    iHwnd = WindowFromPoint(pt.X, pt.Y)
    
    If iHwnd <> UserControl.hWnd Then
        Timer1.Enabled = False
        m_MouseIn = False
        Call DrawButton
    End If
    
End Sub

Private Sub UserControl_EnterFocus()
    m_GotFocus = True
    Call DrawButton
End Sub

Private Sub UserControl_ExitFocus()
    m_GotFocus = False
    Call DrawButton
End Sub

Private Sub UserControl_InitProperties()
    BackGround = &HF0F0F0
    BackgroundHover = &HF4D8B2
    TextColor = vbBlack
    TextColorHover = vbWhite
    BorderHover = &HD77800
    MouseDownBackground = &HD4BB9A
    TextColorDown = vbBlack
    Text = Ambient.DisplayName
    ShowFocusRect = False
    TextAlign = taCenter
    PictureAlign = iaCenter
    Display = TextOnly
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseButton = Button
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbLeftButton Then
        m_MouseDown = True
        Call DrawButton
     End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If (X >= 0 And Y >= 0) And (X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight) Then
        Timer1.Enabled = True
        m_MouseIn = True
        Call DrawButton
    End If
End Sub

Public Property Get BackGround() As OLE_COLOR
    BackGround = m_BackColorDefault
End Property

Public Property Let BackGround(ByVal vNewValue As OLE_COLOR)
    m_BackColorDefault = vNewValue
    Call DrawButton
    PropertyChanged "BackGround"
End Property

Public Property Get BackgroundHover() As OLE_COLOR
    BackgroundHover = m_BackColorHover
End Property

Public Property Let BackgroundHover(ByVal vNewValue As OLE_COLOR)
    m_BackColorHover = vNewValue
    Call DrawButton
    PropertyChanged "BackgroundHover"
End Property

Public Property Get TextColorHover() As OLE_COLOR
    TextColorHover = m_HoverTextColor
End Property

Public Property Let TextColorHover(ByVal vNewValue As OLE_COLOR)
    m_HoverTextColor = vNewValue
    Call DrawButton
    PropertyChanged "TextColorHover"
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal vNewValue As OLE_COLOR)
    m_TextColor = vNewValue
    Call DrawButton
    PropertyChanged "TextColor"
End Property

Public Property Get BorderHover() As OLE_COLOR
    BorderHover = m_BorderHover
End Property

Public Property Let BorderHover(ByVal vNewValue As OLE_COLOR)
    m_BorderHover = vNewValue
    Call DrawButton
    PropertyChanged "BorderHover"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    m_MouseDown = False
    Call DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackGround = PropBag.ReadProperty("BackGround", &HF0F0F0)
    BackgroundHover = PropBag.ReadProperty("BackgroundHover", &HF4D8B2)
    TextColor = PropBag.ReadProperty("TextColor", vbBlack)
    TextColorHover = PropBag.ReadProperty("TextColorHover", vbWhite)
    BorderHover = PropBag.ReadProperty("BorderHover", &HD77800)
    Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
    MouseDownBackground = PropBag.ReadProperty("MouseDownBackground", &HD4BB9A)
    ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", False)
    TextAlign = PropBag.ReadProperty("TextAlign", taCenter)
    PictureAlign = PropBag.ReadProperty("PictureAlign", iaCenter)
    Display = PropBag.ReadProperty("Display", TextOnly)
    TextColorDown = PropBag.ReadProperty("TextColorDown", vbBlack)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

Private Sub UserControl_Resize()
    Call DrawButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackGround", m_BackColorDefault, &HF0F0F0)
    Call PropBag.WriteProperty("BackgroundHover", m_BackColorHover, &HF4D8B2)
    Call PropBag.WriteProperty("TextColor", m_TextColor, vbBlack)
    Call PropBag.WriteProperty("TextColorHover", m_HoverTextColor, vbWhite)
    Call PropBag.WriteProperty("BorderHover", m_BorderHover, &HD77800)
    Call PropBag.WriteProperty("Text", m_Text, Ambient.DisplayName)
    Call PropBag.WriteProperty("MouseDownBackground", m_BackgroundMouseDown, &HD4BB9A)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocus, False)
    Call PropBag.WriteProperty("TextAlign", m_TxtAlign, iaCenter)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PictureAlign", m_PictureAlign, iaCenter)
    Call PropBag.WriteProperty("Display", m_Display, TextOnly)
    Call PropBag.WriteProperty("TextColorDown", m_TextColorDown, vbBlack)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    m_Text = vNewValue
    Call DrawButton
    PropertyChanged "Text"
End Property

Public Property Get MouseDownBackground() As OLE_COLOR
    MouseDownBackground = m_BackgroundMouseDown
End Property

Public Property Let MouseDownBackground(ByVal vNewValue As OLE_COLOR)
    m_BackgroundMouseDown = vNewValue
    Call DrawButton
    PropertyChanged "MouseDownBackground"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocus
End Property

Public Property Let ShowFocusRect(ByVal vNewValue As Boolean)
    m_ShowFocus = vNewValue
    Call DrawButton
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PicImg.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicImg.Picture = New_Picture
    Call DrawButton
    PropertyChanged "Picture"
End Property

Public Property Get TextAlign() As TTextAlign
    TextAlign = m_TxtAlign
End Property

Public Property Let TextAlign(ByVal vNewValue As TTextAlign)
    m_TxtAlign = vNewValue
    Call DrawButton
    PropertyChanged "TextAlign"
End Property

Public Property Get PictureAlign() As TPictureAlign
    PictureAlign = m_PictureAlign
End Property

Public Property Let PictureAlign(ByVal vNewValue As TPictureAlign)
    m_PictureAlign = vNewValue
    Call DrawButton
    PropertyChanged "PictureAlign"
End Property

Public Property Get Display() As TDisplay
    Display = m_Display
End Property

Public Property Let Display(ByVal vNewValue As TDisplay)
    m_Display = vNewValue
    Call DrawButton
    PropertyChanged "Display"
End Property

Public Property Get TextColorDown() As OLE_COLOR
    TextColorDown = m_TextColorDown
End Property

Public Property Let TextColorDown(ByVal vNewValue As OLE_COLOR)
    m_TextColorDown = vNewValue
    Call DrawButton
    PropertyChanged "TextColorDown"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_Click()
    If mMouseButton = vbLeftButton Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    If mMouseButton = vbLeftButton Then
        RaiseEvent DblClick
    End If
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

