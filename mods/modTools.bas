Attribute VB_Name = "modTools"
Option Explicit

'Browse for folder
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public ExportType As Integer    '1 = BMP 2 = TGA
Public ExportOption As Byte
Public Button_Press As Integer  '0 = Cancel 1 = OK

Function FixPath(lPath As String) As String
    If Right(lPath, 1) <> "\" Then
        FixPath = lPath & "\"
    Else
        FixPath = lPath
    End If
End Function

Function GetFilename(lzFile As String) As String
Dim e_pos As Integer
    e_pos = InStrRev(lzFile, "\", Len(lzFile), vbBinaryCompare)
    
    If (e_pos > 0) Then
        GetFilename = Mid(lzFile, e_pos + 1)
    Else
        GetFilename = lzFile
    End If

End Function

Public Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String)
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        Offset = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

