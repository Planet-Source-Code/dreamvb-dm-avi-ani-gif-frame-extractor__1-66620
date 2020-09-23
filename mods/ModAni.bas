Attribute VB_Name = "ModAni"
Option Explicit

Private Type TAniHeader
    RIFF As String * 4
    dwFileSize As Long
    ACON As String * 4
    LIST As String * 4
End Type

Private Type TIconInfo
    ID As String * 4
    dwSize As Long
End Type

Private Type AniInfo
    cFrames As Long
    Title As String
    Credits As String
    Filesize As Long
    wHeight As Integer
    wWidth As Integer
    PixelCount As Integer
    StrFormat As String
End Type

Private AniPtr As Long
Private AniHeader As TAniHeader
Private AniIcon As TIconInfo
Public AniInfo As AniInfo
Private AniFrameData() As String
Private aFilePos As Long
Public AniErr(4) As String

Private Sub AniInitErros()
    'Fill in the error codes array
    AniErr(0) = "File is empty."
    AniErr(1) = "A vaild RIFF ID Was Not Found."
    AniErr(2) = "ACON ID Was Not Found."
    AniErr(3) = "LIST ID Was nNot Found."
    AniErr(4) = "NO frames were found in this file."
End Sub

Function AniGetString(cTag As String) As String
Dim StrTemp As String
Dim iVal As Long
Dim OldPos As Long
    
    OldPos = Seek(AniPtr)
    
    StrTemp = Space(Len(cTag))
    Do Until UCase(StrTemp) = cTag
        'If our Filepos goes past the the file length reset file pos and get out of here
        If (aFilePos > LOF(AniPtr)) Then
            Seek #AniPtr, OldPos
            Exit Function
        End If
        
        aFilePos = aFilePos + 1
        Get #AniPtr, aFilePos, StrTemp
        DoEvents
    Loop
    '
    Get #AniPtr, , iVal
    
    If (iVal) = 0 Then Exit Function
    StrTemp = Space(iVal)
    Get #AniPtr, , StrTemp
    If Right(StrTemp, 1) = Chr(0) Then StrTemp = Left(StrTemp, Len(StrTemp) - 1)
    AniGetString = StrTemp
    'Clean Up
    iVal = 0
    StrTemp = ""
    
End Function

Public Sub CleanUp()
    'Clear up any used variables
    ZeroMemory AniHeader, Len(AniHeader)
    ZeroMemory AniIcon, Len(AniIcon)
    ZeroMemory AniInfo, Len(AniInfo)
    aFilePos = 1
End Sub

Function ExtractAniFrame(lzFile As String, Index As Integer, SaveFile As String, Optional ResType As Integer = &H1) As Boolean
Dim vData As Variant
Dim sData() As Byte
Dim fin As Long, fout As Long
On Error Resume Next
    
    If (Index < 0) Or (Index > UBound(AniFrameData)) Then Exit Function
    'Split the icon data offset and data len
    vData = Split(AniFrameData(Index), Chr(0))
    
    fin = FreeFile
    Open lzFile For Binary As #fin
        'Check filesize
        If LOF(fin) = 0 Then Exit Function
        'Resize array to hold the data
        ReDim sData(0 To CLng(vData(1)))
        'Move to the data offset
        Seek #fin, CLng(vData(0))
        'Extract the data
        Get #fin, CLng(vData(0)), sData
    Close #fin
    
    sData(2) = ResType  'Resource type icon = 1 cursor = 2
    sData(8) = &H10     'I am not to sure.
    'Padd out with some chr(0)'s
    sData(9) = &H0
    sData(10) = &H0
    sData(11) = &H0
    
    'Save the new icon
    fout = FreeFile
    Open SaveFile For Binary As #fout
        Put #fout, , sData
    Close #fout
    
    'Clear up
    Erase vData
    Erase sData
    ExtractAniFrame = True
    
End Function

Function OpenAniFile(lzFile As String) As Integer
Dim FrameTag As String * 4
Dim x As Long
Dim AniFrame As Long
Dim dwOffset As Long
Dim StrA As String, StrB As String
Dim TmpData() As Byte
Dim OldPos As Long

    AniPtr = FreeFile
    OpenAniFile = -1
    Call AniInitErros
    
    Open lzFile For Binary As #AniPtr
        Get #AniPtr, , AniHeader
        
        If (AniHeader.RIFF <> "RIFF") Then
            OpenAniFile = 1
            Exit Function
        ElseIf (AniHeader.ACON <> "ACON") Then
            OpenAniFile = 2
            Exit Function
        'ElseIf (AniHeader.LIST <> "LIST") Then
            'OpenAniFile = 3
            'Exit Function
        Else
            aFilePos = Seek(AniPtr) + 1
            StrA = Trim(AniGetString("INAM")) 'Extract Title Text
            StrB = Trim(AniGetString("IART")) 'Extract Credits Text
        End If
        
        aFilePos = Seek(AniPtr) + 1
        'Locate the frame tag frame, this is were the icons data starts
        Do
            aFilePos = aFilePos + 1
            Get #AniPtr, aFilePos, FrameTag
            DoEvents
        Loop Until (FrameTag = "fram")
        
        'Loop to the end of the file and extract each single icon from the cursor
        For x = aFilePos + 1 To LOF(AniPtr)
            'Get the icon data
            Get #AniPtr, x, AniIcon
            'Check that we have the icon tag
            If (AniIcon.ID = "icon") Then
                'Ok here we just want to get the first frames data
                ' so we can get some information about the icon
                If (AniFrame = 0) Then
                    'Keep the old file position
                    OldPos = Seek(AniPtr)
                    ReDim TmpData(AniIcon.dwSize - 1) As Byte
                    'Get the data
                    Get #AniPtr, (OldPos - Len(AniIcon) + 8), TmpData
                    'Now seek back to our old filepos
                    Seek #AniPtr, OldPos
                End If
                'Get current file position
                aFilePos = Seek(AniPtr)
                'Below is the offset of the icon found
                dwOffset = (aFilePos - Len(AniIcon) + 8)
                ReDim Preserve AniFrameData(AniFrame) As String
                AniFrameData(AniFrame) = dwOffset & Chr(0) & (AniIcon.dwSize - 1)
                AniFrame = AniFrame + 1
            End If
        Next
        'Clear up used variables
        'Icon with and height are located in byte 6 and 7 and pixel count is 36
        'Fill Icon for this resource
        With AniInfo
            .cFrames = AniFrame
            .Filesize = AniHeader.dwFileSize
            .Credits = StrB
            .Title = StrA
            .wWidth = TmpData(6)
            .wHeight = TmpData(7)
            .PixelCount = TmpData(36)
            
            If (.PixelCount = 4) Then
                .StrFormat = "4 bit, 16 colors"
            ElseIf (.PixelCount = 8) Then
                .StrFormat = "8 bit, 256 colors"
            ElseIf (.PixelCount = 24) Then
                .StrFormat = "24 bit, true color"
            ElseIf (.PixelCount = 16) Then
                .StrFormat = "16 bit, true color"
            Else
                .StrFormat = ""
            End If
            
            
        End With
        
        
        
        Erase TmpData
        x = 0
        AniFrame = 0
        dwOffset = 0
        OldPos = 0
        StrA = ""
        StrB = ""
    Close #AniPtr
End Function

