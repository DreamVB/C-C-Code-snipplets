VERSION 5.00
Begin VB.UserControl dmDlg32 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   405
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Picture         =   "dmDlg32.ctx":0000
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "dmDlg32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As TOPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As TOPENFILENAME) As Long

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type TOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Enum TShowDialog
    dlgOPEN = 0
    dlgSAVE = 1
    dlgColor = 2
End Enum

Enum TFlagTypes
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
    'Color dialog
    CC_RGBINIT = &H1&
    CC_FULLOPEN = &H2&
    CC_PREVENTFULLOPEN = &H4&
    CC_SOLIDCOLOR = &H80&
    CC_ANYCOLOR = &H100&
    CLR_INVALID = &HFFFF
End Enum

Private Const MAX_PATH As Integer = 260

Private m_Filename As String
Private m_FileTitle As String
Private m_Filter As String
Private m_InitialDir As String
Private m_Title As String
Private m_DefaultExt As String
Private m_Color As OLE_COLOR
Private m_Flags As TFlagTypes
Private m_ShowDlg As TShowDialog
Private od As TOPENFILENAME
Private cd As CHOOSECOLOR

Private Function CharAt(Start As Integer, S As String, FindChar As String) As Integer
Dim I As Integer
    'Get position of char in string
    For I = Start To Len(S)
        If Mid$(S, I, 1) = FindChar Then
            Exit For
        End If
    Next I
    CharAt = I
End Function

Private Function ReplaceChar(S As String, FindChar As String, ReplaceWith As String)
Dim I As Integer
Dim ch As String
Dim Buffer As String
    'Replace a char in a string
    For I = 1 To Len(S)
        'Get single char
        ch = Mid$(S, I, 1)
        'Look for the FindChar
        If ch = FindChar Then
            'Replace current char with ReplaceWith
            ch = ReplaceWith
        End If
        
        Buffer = Buffer & ch
    Next I
    
    ReplaceChar = Buffer
End Function

Private Function StripNull(S As String) As String
Dim xPos As Integer
    'Strip char 0 from string
    xPos = CharAt(1, S, Chr$(0))
    
    If xPos <> 0 Then
        StripNull = Left$(S, xPos)
    Else
        StripNull = S
    End If
End Function

Public Function GetFilterIndex() As Integer
    'Return the user selected filter index
    GetFilterIndex = od.nFilterIndex
End Function

Public Function Execute() As Boolean
Dim Ret As Long
ReDim CustomColors(0 To 16 * 4 - 1) As Byte
Dim I As Integer
    
    'Zero out types data
    Call ZeroMemory(od, Len(od))
    Call ZeroMemory(cd, Len(cd))
    'Check the displaying dialof to be shown
    If ShowDialogType = dlgOPEN Or ShowDialogType = dlgSAVE Then
        'od is used for the open and save dialogs
        With od
            .lStructSize = Len(od)
            .hwndOwner = Parent.hWnd
            .hInstance = App.hInstance
            .lpstrInitialDir = InitialDir
            .lpstrFilter = ReplaceChar(DlgFilter, "|", Chr(0))
            .lpstrFile = Space$(MAX_PATH - 1)
            .nMaxFile = MAX_PATH
            .lpstrFileTitle = Space(MAX_PATH - 1)
            .nMaxFileTitle = MAX_PATH
            .lpstrTitle = Title
            .lpstrDefExt = DefaultExt
        End With
        
        If ShowDialogType = dlgOPEN Then
            Execute = CBool(GetOpenFileName(od))
        End If
    
        If ShowDialogType = dlgSAVE Then
            Execute = CBool(GetSaveFileName(od))
        End If
        
        'If dialog executed then set Filename and Filetitle
        If Execute <> 0 Then
            'Get the string data upto the nullchar
            Filename = StripNull(od.lpstrFile)
            FileTitle = StripNull(od.lpstrFileTitle)
        End If
    End If
    'Show color dialog
    If ShowDialogType = dlgColor Then
        With cd
            'Set custom colors
            For I = LBound(CustomColors) To UBound(CustomColors)
                CustomColors(I) = 0
            Next I
            'Fill in color type data
            .lStructSize = Len(cd)
            .hInstance = App.hInstance
            .rgbResult = Color
            .hwndOwner = Parent.hWnd
            .lpCustColors = StrConv(CustomColors, vbUnicode)
            .flags = flags
            .rgbResult = Color
            'Show and return color dialog result
            Execute = CBool(CHOOSECOLOR(cd))
            
            If Execute Then
                'Set the return color property
                Color = .rgbResult
            End If
        End With
    End If
End Function

Public Property Get Filename() As String
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal vNewValue As String)
    m_Filename = vNewValue
    PropertyChanged "Filename"
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal vNewValue As String)
    m_Title = vNewValue
    PropertyChanged "Title"
End Property

Public Property Get DlgFilter() As String
    DlgFilter = m_Filter
    PropertyChanged "DlgFilter"
End Property

Public Property Let DlgFilter(ByVal vNewValue As String)
    m_Filter = vNewValue
End Property

Public Property Get InitialDir() As String
    InitialDir = m_InitialDir
End Property

Public Property Let InitialDir(ByVal vNewValue As String)
    m_InitialDir = vNewValue
    PropertyChanged "InitialDir"
End Property

Private Sub UserControl_InitProperties()
    Filename = ""
    Title = ""
    DlgFilter = ""
    DefaultExt = ""
    InitialDir = "C:\"
    flags = 0
    Color = vbBlack
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Filename = PropBag.ReadProperty("Filename", "")
    Title = PropBag.ReadProperty("Title", "")
    FileTitle = PropBag.ReadProperty("FileTitle", "")
    DlgFilter = PropBag.ReadProperty("DlgFilter", "")
    InitialDir = PropBag.ReadProperty("InitialDir", "C:\")
    ShowDialogType = PropBag.ReadProperty("ShowDialogType", dlgOPEN)
    flags = PropBag.ReadProperty("Flags", 0)
    DefaultExt = PropBag.ReadProperty("DefaultExt", "")
    Color = PropBag.ReadProperty("Color", vbBlack)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Width = 405
    UserControl.Height = 405
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Filename", m_Filename, "")
    Call PropBag.WriteProperty("Title", m_Title, "")
    Call PropBag.WriteProperty("FileTitle", m_FileTitle, "")
    Call PropBag.WriteProperty("DlgFilter", m_Filter, "")
    Call PropBag.WriteProperty("InitialDir", m_InitialDir, "C:\")
    Call PropBag.WriteProperty("ShowDialogType", m_ShowDlg, dlgOPEN)
    Call PropBag.WriteProperty("Flags", m_Flags, 0)
    Call PropBag.WriteProperty("DefaultExt", m_DefaultExt, "")
    Call PropBag.WriteProperty("Color", m_Color, vbBlack)
End Sub
 
Public Property Get FileTitle() As String
    FileTitle = m_FileTitle
End Property

Public Property Let FileTitle(ByVal vNewValue As String)
    m_FileTitle = vNewValue
    PropertyChanged "FileTitle"
End Property

Public Property Get ShowDialogType() As TShowDialog
    ShowDialogType = m_ShowDlg
End Property

Public Property Let ShowDialogType(ByVal vNewValue As TShowDialog)
    m_ShowDlg = vNewValue
    PropertyChanged "ShowDialogType"
End Property

Public Property Get flags() As TFlagTypes
    flags = m_Flags
End Property

Public Property Let flags(ByVal vNewValue As TFlagTypes)
    m_Flags = vNewValue
    PropertyChanged "Flags"
End Property

Public Property Get DefaultExt() As String
    DefaultExt = m_DefaultExt
End Property

Public Property Let DefaultExt(ByVal vNewValue As String)
    m_DefaultExt = vNewValue
    PropertyChanged "DefaultExt"
End Property

Public Property Get Color() As OLE_COLOR
    Color = m_Color
End Property

Public Property Let Color(ByVal vNewValue As OLE_COLOR)
    m_Color = vNewValue
    PropertyChanged "Color"
End Property
