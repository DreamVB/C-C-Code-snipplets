Attribute VB_Name = "modVBHelper"
' DM Helper mod for Visual Basic
' by dreamvb
' Version 1.0
' comments or questions to dreamvb@outlook.com
' also find on https://github.com/DreamVB

Option Explicit

Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Enum TFileCase
    aUPPER = 0
    fLower = 1
End Enum


Public Const MAX_PATH As Integer = 260

Public Sub AppendString(Dest As String, Src As String)
    Dest = Dest & Src
End Sub

Public Function SameString(S1 As String, S2 As String)
    SameString = (S1 = S2)
End Function

Public Function IsUpperCase(S1 As String)
    IsUpperCase = S1 = UCase$(S1)
End Function

Public Function IsLowerCase(S1 As String) As Boolean
    IsLowerCase = S1 = LCase$(S1)
End Function

Public Function IsDigit(S1 As String) As Boolean
    IsDigit = (S1 Like "[0-9]")
End Function

Public Function IsAlpha(S1 As String) As Boolean
Dim ch As String
    If Len(S1) = 0 Then Exit Function
    ch = UCase$(Left$(S1, 1))
    IsAlpha = (ch Like "[A-Z]")
End Function

Function IsWhiteSpace(S1 As String) As Boolean
Dim ch As String * 1
    
    ch = S1
    
    If ch = " " Or ch = vbLf Or _
    ch = vbCr Or ch = vbTab Then
        IsWhiteSpace = True
    End If
    
End Function

Function IsStrEmpty(S1 As String) As Boolean
    IsStrEmpty = Len(Trim$(S1)) = 0
End Function

Public Function IsAlphaDigit(S1 As String) As Boolean
Dim X As Integer
Dim Vaild As Boolean
Dim ch As String * 1

    Vaild = True
    
    For X = 1 To Len(S1)
        ch = Mid$(S1, X, 1)
        If Not IsAlpha(ch) And _
        IsDigit(ch) = False Then
            Exit Function
        End If
    Next X
    
    IsAlphaDigit = Vaild
    
End Function

Public Function CountChars(S1 As String, Char As String) As Integer
Dim X As Integer
Dim nChars As Integer
    
    nChars = 0
    For X = 1 To Len(S1)
        If Mid$(S1, X, 1) = Char Then
            nChars = (nChars + 1)
        End If
    Next X
    CountChars = nChars
End Function

Public Function CountStrings(S1 As String, sString As String, Delimiter As String) As Integer
Dim X As Integer
Dim StrS As String
Dim Temp As String
Dim nCount As Integer
Dim ch As String * 1

    nCount = 0
    Temp = S1
    
    If Not EndsWith(Temp, " ") Then
        Call AppendString(Temp, " ")
    End If
    
    For X = 1 To Len(Temp)
        ch = Mid$(Temp, X, 1)
        
        If ch = Delimiter Then
            If Trim$(StrS) = sString Then
                nCount = (nCount + 1)
            End If
            StrS = ""
        End If
        StrS = StrS & ch
    Next X
    
    CountStrings = nCount
    
End Function

Public Function CompareFileName(S1 As String, S2 As String) As Boolean
    CompareFileName = LCase$(S1) = LCase$(S2)
End Function

Public Function SameFileName(S1 As String, S2 As String)
    SameFileName = (S1 = S2)
End Function

Public Function IsLeapYear(m_year As Integer) As Boolean
  IsLeapYear = (m_year Mod 4 = 0) And ((m_year Mod 100 <> 0) Or (m_year Mod 400 = 0))
End Function

Public Function QuoteString(S1 As String) As String
    QuoteString = Chr$(34) & S1 & Chr(34)
End Function

Public Function FileNameCase(Filename, FileCase As TFileCase)
    If FileCase = aUPPER Then
        FileNameCase = UCase$(Filename)
    End If
    If FileCase = fLower Then
        FileNameCase = LCase$(Filename)
    End If
End Function

Public Function ExtractQuoteString(S1 As String)
    If Left$(S1, 1) = Chr$(34) Then
        If Right$(S1, 1) = Chr$(34) Then
            ExtractQuoteString = Mid$(S1, 2, Len(S1) - 2)
        Else
            ExtractQuoteString = S1
        End If
        
    Else
        ExtractQuoteString = S1
    End If
End Function

Public Function StartsWith(S1 As String, S2 As String) As Boolean
    StartsWith = Left(S1, Len(S2)) = S2
End Function

Public Function EndsWith(S1 As String, S2 As String) As Boolean
    EndsWith = Right(S1, Len(S2)) = S2
End Function

Public Function FirstChar(S1 As String) As String
    If Len(S1) <> 0 Then
        FirstChar = Left$(S1, 1)
    End If
End Function

Public Function LastDelimiter(S1 As String, Delimiter As String)
Dim I As Integer
Dim nIdx As Integer
    
    nIdx = -1

    For I = 1 To Len(S1)
        If Mid$(S1, I, 1) = Delimiter Then
            nIdx = I
        End If
    Next I
    
    LastDelimiter = nIdx
    
End Function

Public Function ChangeFileExt(Filename As String, Extension As String)
Dim X As Integer
    X = LastDelimiter(Filename, ".")
    If X <> -1 Then
        ChangeFileExt = Left$(Filename, X - 1) & Extension
    End If
End Function

Public Function GetFileExtension(Filename As String) As String
Dim X As Integer
    X = LastDelimiter(Filename, ".")
    If X <> -1 Then
        GetFileExtension = Mid$(Filename, X)
    End If
End Function

Public Function ExtractFilename(Filename As String) As String
Dim X As Integer
    X = LastDelimiter(Filename, "\")
    If X <> -1 Then
        ExtractFilename = Mid$(Filename, X + 1)
    End If
End Function

Public Function ExtractDrive(Filename As String) As String
Dim X As Integer
    X = CharAt(Filename, ":")
    If X <> -1 Then
        ExtractDrive = Left$(Filename, X)
    End If
End Function

Public Function ExtractFilePath(Filename As String) As String
Dim X As Integer
    X = LastDelimiter(Filename, "\")
    If X <> -1 Then
        ExtractFilePath = Left$(Filename, X)
    End If
End Function

Public Function ExtractFilePathNoDrive(Filename As String) As String
Dim lzPath As String
Dim X As Integer
   
    lzPath = ExtractFilePath(Filename)
    X = CharAt(lzPath, "\")
    lzPath = Mid$(Filename, X)
    X = LastDelimiter(lzPath, "\")
    
    If X <> -1 Then
        ExtractFilePathNoDrive = Mid$(lzPath, 1, X)
    End If
        
End Function

Public Function ExtractShortPathName(Filename As String) As String
Dim Buffer As String
Dim Ret As Long
    Buffer = Space$(MAX_PATH)
    
    Ret = GetShortPathName(Filename, Buffer, MAX_PATH)
    
    If Ret <> 0 Then
        ExtractShortPathName = Left(Buffer, Ret)
    End If
    
End Function

Public Function ChangeFilePath(S1 As String, NewPath As String) As String
    ChangeFilePath = FixPath(NewPath) & ExtractFilename(S1)
End Function

Public Function ExpandFileName(Filename As String) As String
Dim Buffer As String
Dim Ret As Long
    Buffer = Space(MAX_PATH)
    
    Ret = GetFullPathName(Filename, MAX_PATH, Buffer, "")
    
    If Ret <> 0 Then
        ExpandFileName = Left$(Buffer, CharAt(Buffer, Chr(0)))
    End If
    
    Buffer = vbNullString
End Function

Public Function FileExists(Filename As String) As Boolean
Dim wfd As WIN32_FIND_DATA
Dim Ret As Long
    
    Ret = FindFirstFile(Filename, wfd)
    If Ret <> -1 Then
        FileExists = True
    Else
        FileExists = False
    End If
    
    Ret = FindClose(Ret)
    
End Function

Public Function FileIsReadOnly(Filename As String) As Boolean
    If FileExists(Filename) <> False Then
        FileIsReadOnly = GetFileAttributes(Filename) And &H1
    End If
End Function

Public Function DirExists(Filename As String) As Boolean
Dim fAttr As Long
    
    fAttr = GetFileAttributes(Filename)
    If fAttr <> -1 Then
        DirExists = (fAttr = 16)
    End If
    
End Function

Public Function FixPath(S1 As String) As String
    If Right$(S1, 1) <> "\" Then
        FixPath = S1 & "\"
    Else
        FixPath = S1
    End If
End Function

Public Function LastChar(S1 As String) As String
    If Len(S1) <> 0 Then
        LastChar = Right$(S1, 1)
    End If
End Function

Public Function CharAt(S1 As String, ch As String)
Dim I As Integer
    For I = 1 To Len(S1)
        If Mid$(S1, I, 1) = ch Then
            Exit For
        End If
    Next I
    CharAt = I
End Function
