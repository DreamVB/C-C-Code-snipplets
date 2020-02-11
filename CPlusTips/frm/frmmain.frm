VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DM C++ Code Reader"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   9420
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   StartUpPosition =   2  'CenterScreen
   Begin Project1.dmDlg32 dmDlg321 
      Left            =   270
      Top             =   5250
      _ExtentX        =   714
      _ExtentY        =   714
      InitialDir      =   ""
      ShowDialogType  =   1
   End
   Begin VB.PictureBox pFilterAlpha 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   628
      TabIndex        =   12
      Top             =   1260
      Width           =   9420
      Begin VB.Label lblFilter 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Filter:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   14
         Top             =   60
         Width           =   675
      End
      Begin VB.Label lblAlpha 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Index           =   0
         Left            =   885
         MouseIcon       =   "frmmain.frx":0E9A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   60
         Width           =   270
      End
   End
   Begin VB.PictureBox PicCats 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00928B80&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   628
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   690
      Width           =   9420
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5415
         TabIndex        =   16
         Top             =   105
         Width           =   3960
      End
      Begin Project1.Flat Flat1 
         Height          =   390
         Left            =   1650
         TabIndex        =   11
         Top             =   75
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   688
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
         BorderColor     =   -2147483643
      End
      Begin VB.ComboBox cboCats 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   3105
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4845
         TabIndex        =   15
         Top             =   105
         Width           =   495
      End
      Begin VB.Label lblCats 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categories:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   195
         TabIndex        =   9
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.ListBox lstCodes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3030
      IntegralHeight  =   0   'False
      Left            =   15
      TabIndex        =   0
      Top             =   1725
      Width           =   9360
   End
   Begin VB.PictureBox pSpacer 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E3C9AE&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   628
      TabIndex        =   3
      Top             =   4830
      Width           =   9420
   End
   Begin Project1.dStatusbar dStatusbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   4845
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BackColor       =   7630178
      GripStyle       =   1
   End
   Begin VB.PictureBox pBanner 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00746D62&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   628
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9420
      Begin Project1.dmFlatButton cmdView 
         Height          =   555
         Left            =   225
         TabIndex        =   4
         ToolTipText     =   "Source View"
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   979
         BackGround      =   7630178
         Text            =   "dmFlatButton1"
         Picture         =   "frmmain.frx":0FEC
         Display         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.dmFlatButton cmdSave 
         Height          =   555
         Left            =   870
         TabIndex        =   5
         ToolTipText     =   "Export"
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   979
         BackGround      =   7630178
         Text            =   "dmFlatButton1"
         Picture         =   "frmmain.frx":1C3E
         Display         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.dmFlatButton cmdClose 
         Height          =   555
         Left            =   2970
         TabIndex        =   6
         ToolTipText     =   "Close"
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   979
         BackGround      =   7630178
         Text            =   "dmFlatButton1"
         Picture         =   "frmmain.frx":2890
         Display         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.dmFlatButton cmdAbout 
         Height          =   555
         Left            =   2250
         TabIndex        =   7
         ToolTipText     =   "About"
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   979
         BackGround      =   7630178
         Text            =   "About"
         Picture         =   "frmmain.frx":34E2
         Display         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.dmFlatButton cmdCopy 
         Height          =   555
         Left            =   1485
         TabIndex        =   17
         ToolTipText     =   "Copy Sourcecode"
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   979
         BackGround      =   7630178
         Text            =   "dmFlatButton1"
         Picture         =   "frmmain.frx":4134
         Display         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&HELP"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TBagHead
    Sig As String * 3
    LastUpdate As String
    NOF As Long
End Type

Private Type TBag
    Filename As String
    FileData As String
End Type

Private Files() As TBag
Private BagHeader As TBagHead
Private m_TipsPackage As String
Private m_DataPath As String
Private m_Button As MouseButtonConstants
Const FILTER_TEXT_COLOR As Long = 1152750

Private Sub UpdateStatusbar()
    'Set statusbar text
    dStatusbar1.SimpleText = "Tips(s) " & CStr(lstCodes.ListCount) & "  |  Last-Update " & BagHeader.LastUpdate
End Sub

Private Sub SetFilterIndexColor(Index As Integer)
Dim I As Integer
    For I = 0 To lblAlpha.Count - 1
        lblAlpha(I).ForeColor = FILTER_TEXT_COLOR
    Next I
    lblAlpha(Index).ForeColor = vbRed
End Sub

Private Sub SaveSourceFile(Filename As String, sData As String)
Dim fp As Long
    fp = FreeFile
    Open Filename For Binary As #fp
        Put #fp, , sData
    Close #fp
End Sub

Private Sub FilterCodes(FilterText As String)
Dim Idx As Long
Dim I As Long
Dim sItem As String
    
    Idx = 0
    'Clear listbox
    Call lstCodes.Clear
    
    'Load the tips names into the listbox.
    For I = 0 To UBound(Files)
        'Check if the user has filtered by letter
        If Len(FilterText) <> 0 Then
            If UCase$(Left$(Files(I).Filename, 1)) = FilterText Then
                'INC index
                Idx = (Idx + 1)
                'Padd the tip name with a number and filename
                sItem = Files(I).Filename
                'Add item to listbox
                Call lstCodes.AddItem(sItem)
                'Set listbox item data
                lstCodes.ItemData(lstCodes.NewIndex) = I
            End If
        Else
            sItem = Files(I).Filename
            'Add item to listbox
            Call lstCodes.AddItem(sItem)
            'Set listbox item data
            lstCodes.ItemData(lstCodes.NewIndex) = I
        End If
    Next I
    'Set statusbar text
    Call UpdateStatusbar
End Sub

Private Sub LoadFilterIndex()
Dim I As Integer
    For I = 1 To 25
        'Load a new label control
        Load lblAlpha(I)
        'Position the label
        lblAlpha(I).Left = lblAlpha(I - 1).Left + lblAlpha(I).Width
        'Set label caption
        lblAlpha(I).Caption = Chr$(64 + I)
        'Set label forecolor
        lblAlpha(I).ForeColor = FILTER_TEXT_COLOR
        'Show the lable
        lblAlpha(I).Visible = True
    Next I
End Sub

Private Sub LoadCats()
Dim X As String
Dim xFile As String
    
    X = Dir$(m_DataPath, vbNormal Or vbArchive)
    
    While X <> vbNullString
        'Remove the file ext
        xFile = ChangeFileExt(X, "")
        'Add file title to listbox
        Call cboCats.AddItem(xFile)
        'Get next file.
        X = Dir$()
    Wend
    
End Sub

Private Function Encode(Source As String)
Dim I As Long
Dim ch As Byte
Dim s0 As String

    For I = 1 To Len(Source)
        ch = Asc(Mid$(Source, I, 1)) Xor 128
        Call AppendString(s0, Chr$(ch))
    Next I
    
    Encode = s0
    
End Function

Private Sub LoadPackage(Filename As String)
Dim fp As Long
    'Get free file
    fp = FreeFile

    Open Filename For Binary As #fp
        Get #fp, , BagHeader
        'Check header
        If (BagHeader.Sig <> "BAG") Then
            Call MsgBox("Invaild Package File.", vbCritical, "#Error_65")
            Exit Sub
        Else
            'Resize the array to hold the file data
            ReDim Preserve Files(0 To BagHeader.NOF - 1) As TBag
            'Load files data
            Get #fp, , Files
        End If
    Close #fp
End Sub

Private Sub cboCats_Click()
    m_TipsPackage = m_DataPath & cboCats.List(cboCats.ListIndex) & ".bag"
    'Load the tips package
    Call LoadPackage(m_TipsPackage)
    Call FilterCodes(vbNullString)
End Sub

Private Sub cmdAbout_Click()
    Call frmabout.Show(vbModal)
End Sub

Private Sub cmdClose_Click()
    Call Unload(frmmain)
End Sub

Private Sub cmdCopy_Click()
On Error Resume Next
    If IsStrEmpty(SourceCode) Then Exit Sub
    Call Clipboard.Clear
    Call Clipboard.SetText(SourceCode)
End Sub

Private Sub cmdSave_Click()
    If IsStrEmpty(SourceCode) Then Exit Sub
    With dmDlg321
        .Title = "Export Sourcecode"
        .DlgFilter = "Text Files(*.txt)|*.txt|C++ Source Files(*.cpp)|*.cpp|C Source Files(*.c)|*.c"
        If .Execute Then
            'Save source code to file
            Call SaveSourceFile(.Filename, SourceCode)
        End If
    End With
End Sub

Private Sub cmdView_Click()
    If IsStrEmpty(SourceCode) Then Exit Sub
    'Show code view form
    Call frmcodeview.Show(vbModal)
End Sub

Private Sub Form_Load()
    lblAlpha(0).ForeColor = vbRed
    Call LoadFilterIndex
    m_DataPath = FixPath(App.Path) & "data\"
    'Load code Categories
    DoEvents
    Call LoadCats
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lstCodes.Height = (frmmain.ScaleHeight - dStatusbar1.Height - lstCodes.Top) - 4
    lstCodes.Width = (frmmain.ScaleWidth - lstCodes.Left) - 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoEvents
    Erase Files
    Set frmmain = Nothing
End Sub

Private Sub lblAlpha_Click(Index As Integer)
    
    If (m_Button = vbLeftButton) Then
        'Exit if no cat is selected
        If cboCats.ListIndex = -1 Then
            Exit Sub
        End If
        'Check for clear filter click
        If lblAlpha(Index).Caption = "#" Then
            Call FilterCodes("")
        Else
            'Filter items by first letter
            Call FilterCodes(lblAlpha(Index).Caption)
        End If
    End If
    
    Call SetFilterIndexColor(Index)
End Sub

Private Sub lblAlpha_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    m_Button = Button
End Sub

Private Sub lblAlpha_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If lblAlpha(Index).Caption = "#" Then
        lblAlpha(Index).ToolTipText = "Clear Filter"
    Else
        lblAlpha(Index).ToolTipText = "Filter by " & lblAlpha(Index).Caption
    End If
End Sub

Private Sub lstCodes_Click()
Dim Id As String
    Id = lstCodes.ItemData(lstCodes.ListIndex)
    SourceCode = Encode(Files(Id).FileData)
End Sub

Private Sub lstCodes_DblClick()
    'Show code view form
    Call frmcodeview.Show(vbModal)
End Sub

Private Sub lstCodes_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) And (lstCodes.ListCount > 0) Then
        Call lstCodes_DblClick
    End If
End Sub

Private Sub mnuAbout_Click()
    Call cmdAbout_Click
End Sub

Private Sub mnuExit_Click()
    Call cmdClose_Click
End Sub

Private Sub pBanner_Resize()
    Call pBanner.Cls
    pBanner.Line (0, pBanner.ScaleHeight - 1)-(pBanner.ScaleWidth - 1, pBanner.ScaleHeight - 1), &HBAB3A8
End Sub

Private Sub PicCats_Resize()
On Error Resume Next
    txtFind.Width = PicCats.ScaleWidth - txtFind.Left - 8
End Sub

Private Sub txtFind_Change()
    If IsStrEmpty(txtFind.Text) Or lstCodes.ListCount = 0 Then
        Exit Sub
    End If
    
    Call modTools.FindListItem(lstCodes, txtFind.Text)
End Sub
