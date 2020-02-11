VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Text Packager"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFilename 
      Height          =   345
      Left            =   285
      TabIndex        =   2
      Text            =   "C:\out\tips\package.pak"
      Top             =   1590
      Width           =   4920
   End
   Begin VB.TextBox txtSource 
      Height          =   345
      Left            =   285
      TabIndex        =   1
      Text            =   "C:\out\tips\"
      Top             =   555
      Width           =   4920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pack Files"
      Height          =   495
      Left            =   330
      TabIndex        =   0
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label lblPackage 
      AutoSize        =   -1  'True
      Caption         =   "Pacakge file"
      Height          =   195
      Left            =   285
      TabIndex        =   4
      Top             =   1275
      Width           =   885
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "Files to package"
      Height          =   195
      Left            =   285
      TabIndex        =   3
      Top             =   210
      Width           =   1170
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private FileCount As Long

Private Function Encode(Source As String)
Dim I As Integer
Dim ch As Integer
Dim s0 As String

    For I = 1 To Len(Source)
        ch = Asc(Mid(Source, I, 1)) Xor 128
        s0 = s0 & Chr$(ch)
    Next I
    
    Encode = s0
    
End Function

Function ExtractFileTitle(Filename As String) As String
Dim s_pos As Integer
Dim lzFile As String

    s_pos = InStrRev(Filename, "\", Len(Filename))

    If s_pos > 0 Then
        lzFile = Mid$(Filename, s_pos + 1)
    Else
        lzFile = Filename
    End If
    
    s_pos = InStr(lzFile, ".")
    
    If s_pos > 0 Then
        ExtractFileTitle = Mid(lzFile, 1, s_pos - 1)
    Else
        ExtractFileTitle -lzFile
    End If
    
End Function

Private Function OpenFile(Filename As String) As String
Dim fp As Long
Dim Buffer As String
    fp = FreeFile
    
    Open Filename For Binary As #fp
        Buffer = Space(LOF(fp))
        Get #fp, , Buffer
    Close #fp
    
    OpenFile = Buffer
End Function

Private Sub PackFiles(Folder As String, Filename As String)
Dim lzFile As String
Dim fp As Long

    lzFile = Dir$(Folder & "*.*")
    
    While lzFile <> ""
        
        ReDim Preserve Files(0 To FileCount) As TBag
        
        Files(FileCount).Filename = StrConv(ExtractFileTitle(lzFile), vbProperCase)
        Files(FileCount).FileData = Encode(OpenFile(Folder & lzFile))
        FileCount = (FileCount + 1)
        
        lzFile = Dir
    Wend
    
    With BagHeader
        .Sig = "BAG"
        .LastUpdate = CStr(Now)
        .NOF = FileCount
    End With
    
    fp = FreeFile
    Open Filename For Binary As #fp
        Put #fp, , BagHeader
        Put #fp, , Files
    Close #fp
    
    Call MsgBox("Done", vbInformation Or vbOKOnly, Caption)
    
    
End Sub

Private Sub Command1_Click()
    FileCount = 0
    Erase Files
    Call PackFiles(txtSource.Text, txtFilename.Text)
End Sub

