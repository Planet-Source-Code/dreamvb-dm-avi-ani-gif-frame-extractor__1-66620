Attribute VB_Name = "ModGif"
Option Explicit

Private Type TGifFrame
    gDelayTime As Integer
    TransColor As Long
    gEreaseMethod As Byte
    gWidth As Integer
    gHeight As Integer
    gTop As Integer
    gLeft As Integer
    ImgLen As Long
    ImgStartPos As Long
    ImgEndPos As Long
End Type

'GIF FileHeader.
Private Type TGifHeader
    SIG As String * 3
    VER As String * 3
    sWidth As Integer
    sHeight As Integer
    BitFields As Byte
    BackColorIdx As Byte
    PixelRatio As Byte
End Type

Private Type TGifHeadrInfo
    Version As String
    FrameCount As Long
    Width As Integer
    Height As Integer
    sAppName As String
    sComment As String
    dColors As Integer
    dRepeat As Integer
End Type

'GIF Image Descriptor of each frame in the GIF File.
Private Type TGifImgDescriptor
    xLeft As Integer
    xTop As Integer
    Width As Integer
    Height As Integer
    Flags As Byte
End Type

Private Type TGifGraphicCTRL
    cbSize  As Byte
    bFlags  As Byte
    nDelayTime  As Integer
    bTransparentColor   As Byte
    bTerminator As Byte
End Type

Private Type TGifLooping
    lFlags  As Integer
    Looping As Integer
    bTerminator As Byte
End Type

Private gHeader As TGifHeader
Private gImgDesc As TGifImgDescriptor
Private gControl As TGifGraphicCTRL
Private gLoop As TGifLooping
Private gFilePos As Long
Private gFilePtr As Long

Public FrameInfo() As TGifFrame
Public TGifHeadInfo As TGifHeadrInfo

Private abCodes(3) As String
Private EraseCodes(3) As String
Private gFileHead As String
Private m_GifFile As String

Public Sub CleanGlobal()
    gFileHead = ""
    m_GifFile = ""
    ZeroMemory TGifHeadInfo, Len(TGifHeadInfo)
    Erase FrameInfo
End Sub

Public Property Get GifAbort(Index As Integer) As String
    GifAbort = abCodes(Index)
End Property

Public Property Get EraseMethod(Index As Byte) As String
    EraseMethod = EraseCodes(Index)
End Property

Private Function GetGifEOF() As Long
Dim TByte As Byte
Dim Pos As Long

    'Function Used to Return the end of the GIF Frame Block when a Terminator is found chr(0)
    Pos = gFilePos + 2
    Do
        TByte = GetByte(Pos)
        Pos = Pos + (TByte + 1)
        TByte = GetByte(Pos)
        'Loop until we hit the Terminator
    Loop Until (TByte = 0)
    
    GetGifEOF = Pos
    
End Function

Private Function GifFrameCount() As Long
Dim x As Long
Dim Count As Integer

Dim gMagic As String
Dim sTag As String * 3

    'This identfies the start block of each frame in the GIF File
    gMagic = Chr(33) + Chr(249) + Chr(4)
    
    For x = gFilePos To LOF(gFilePtr)
        Get #gFilePtr, x, sTag
        'If we incounter the Magic tag add one to our frame count
        If (sTag = gMagic) Then Count = Count + 1
    Next x
    
    GifFrameCount = Count
    Count = 0
    gMagic = ""
End Function

Function GetByte(Pos As Long) As Byte
    'Return a byte from a file at a spicifed position
    Get #gFilePtr, Pos, GetByte
End Function

Public Sub Init()
    'Add Our GOF Error codes
    abCodes(0) = "Not a GIF file, or incorrect version number"
    abCodes(1) = "Global Color pallete was not found."
    abCodes(2) = "No frames were found in the GIF File."
    abCodes(3) = "Image descripter was not found."
    'Removal codes
    
    EraseCodes(0) = "Default"
    EraseCodes(1) = "Leave"
    EraseCodes(2) = "Restore Background"
    EraseCodes(3) = "Restore Previous"
End Sub

Public Function OpenGIF() As Integer
Dim NumColors As Integer
Dim sComment As String
Dim sAppName As String * 11
Dim gFrameCnt As Long
Dim xCnt As Long
Dim gStart As Long
Dim TByte As Byte
    
    xCnt = 0
    OpenGIF = -1
    gFilePtr = FreeFile
    
    Open m_GifFile For Binary As #gFilePtr
        'Get Gif FileHeader Info
        Get #gFilePtr, , gHeader
        
        'Check for vaild GIF Header
        If (UCase(gHeader.SIG) <> "GIF") Then
            OpenGIF = 0
            Exit Function
        End If
        
        'Check for global pallete
        If (gHeader.BitFields And 128) = 0 Then
            'No glibal pallete found
            OpenGIF = 1
            Exit Function
        End If
  
        'Extract the number of colors in the Gif Pallete
        NumColors = 2 ^ ((gHeader.BitFields And 7) + 1)
        
        'Place File Position pointer over past the pallete
        gFilePos = Len(gHeader) + (NumColors * 3)
        Seek #gFilePtr, gFilePos
        
        'Loop while Control chr is found !
        Do While (GetByte(gFilePos + 1) = 33)
            'Get next byte move 2 places
            TByte = GetByte(gFilePos + 2)
            'Look for Application Marker
            If (TByte = 255) Then
                TByte = GetByte(gFilePos + 3)
                If (TByte = 11) Then
                    Get #gFilePtr, , sAppName
                    Get #gFilePtr, , gLoop
                    gLoop.Looping = gLoop.Looping + 1
                    gFilePos = gFilePos + 19
                End If
            'Check for Comment marker
            ElseIf (TByte = 254) Then
                TByte = GetByte(gFilePos + 3)
                gFilePos = gFilePos + TByte + 4
                Do While GetByte(gFilePos) > 0
                    sComment = sComment & Chr(GetByte(gFilePos + 1))
                    gFilePos = gFilePos + GetByte(gFilePos) + 1
                Loop
            Else
                Exit Do
            End If
            DoEvents
        Loop
        
        'Get Gif FileHeader info
        gFileHead = Space(gFilePos)
        Get #gFilePtr, 1, gFileHead
        'Get a count of each single image in the GIF
        gFrameCnt = GifFrameCount
        
        'Check the GIF has frames
        If (gFrameCnt = 0) Then
            OpenGIF = 2
            Exit Function
        Else
            'Reszie FrameInfo to hold the frame info
            ReDim Preserve FrameInfo(gFrameCnt)
        End If
        
        'Reposition back in our file from were we left off
        Seek #gFilePtr, gFilePos

        Do Until (xCnt = gFrameCnt)

            'Start Position of Single GIF Frame
            gStart = gFilePos + 1
            
            If GetByte(gFilePos + 2) = 249 Then
                Get #gFilePtr, , gControl
                'Remove method
                gControl.bFlags = (gControl.bFlags \ (2 ^ 2) And (2 ^ 3 - 1))
                'Fix Delay if 0 set equal to 1
                If (gControl.nDelayTime = 0) Then gControl.nDelayTime = 1
                gFilePos = Seek(gFilePtr) - 1
            End If
            
            'Check and Get Single frame Image Info
            If GetByte(gFilePos + 1) <> 44 Then
                OpenGIF = 3
                Exit Do
            Else
                Get #gFilePtr, , gImgDesc
                'Interlaced GIF Image Found
                If (gImgDesc.Flags = 135) Then
                    'Add Pallete Size to gFilePos
                    gFilePos = (gFilePos + (NumColors * 3) + 10)
                Else
                    'INC gFilePos by 10
                    gFilePos = gFilePos + 10
                End If
                
                'Fill in Array with each of the frames properties
                With FrameInfo(xCnt)
                    .gEreaseMethod = gControl.bFlags
                    .gDelayTime = gControl.nDelayTime
                    .gHeight = gImgDesc.Height
                    .gLeft = gImgDesc.xLeft
                    .gTop = gImgDesc.xTop
                    .gWidth = gImgDesc.Width
                    .TransColor = gControl.bTransparentColor
                    .ImgEndPos = GetGifEOF
                    .ImgStartPos = gStart
                    .ImgLen = Len(gFileHead) + (.ImgEndPos - .ImgStartPos) + 2
                End With
                gFilePos = GetGifEOF
            End If
            xCnt = xCnt + 1
        Loop
        
        'Fill Our GIF HeaderInfo
        With TGifHeadInfo
            .dColors = NumColors
            .FrameCount = gFrameCnt
            .Height = gHeader.sHeight
            .Width = gHeader.sWidth
            .Version = gHeader.SIG & gHeader.VER
            .sAppName = sAppName
            .sComment = sComment
            .dRepeat = gLoop.Looping
        End With
        
'Clear up used variables
Clean:
    Close #gFilePtr
    gFilePos = 0
    TByte = 0
    NumColors = 0
    gStart = 0
    xCnt = 0
    sComment = ""
    gFrameCnt = 0
    'Clean Used Types
    ZeroMemory gHeader, Len(gHeader)
    ZeroMemory gControl, Len(gControl)
    ZeroMemory gLoop, Len(gLoop)
End Function

Public Sub SaveSingleFrame(lzFile As String, StartPos As Long, EndPos)
Dim fp As Long
Dim sData As String

    fp = FreeFile
    Open m_GifFile For Binary As #fp
        Seek #fp, StartPos
        sData = Space(EndPos - StartPos + 1)
        Get #fp, , sData
    Close #fp
    
    'Save the gif
    Open lzFile For Binary As #fp
        Put #fp, , gFileHead
        Put #fp, , sData
        Put #fp, , ";"
    Close #fp
    
    sData = ""
End Sub

Public Property Let GifFilename(ByVal vNewValue As String)
    m_GifFile = vNewValue
End Property
