Attribute VB_Name = "ModAvi"
Option Explicit

Private Type WavHeader
        FormatTag As Integer
        Channels As Integer
        SampleRate As Long
        BytesPerSecond As Long
        BlockAlignment As Integer
        BitsPerSample As Integer
End Type

Private Type WavHeadInfo
    RIFF As String * 4
    BufferLen As Long
    WavID As String * 4
    FmtID As String * 4
    FmtLength As Long
    WavFmtTag As Integer
    Channels As Integer
    SampleRate As Long
    BytePerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    DataID As String * 4
    wDataLen As Long
End Type

'AVI Header information
Private Type AVIHeader
    dwType As String * 4
    dwFileSize As Long
    dwTAG As String * 4
End Type

'Main AVI Information
Private Type AVIInfo
    ValidID As String * 4
    Reserved As Long
    dwMicroSecPerFrame As Long
    dwMaxBytesPerSec As Long
    dwReserved1 As Long
    dwFlags As Long
    dwTotalFrames As Long
    dwInitialFrame As Long
    dwStreams As Long
    dwBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
End Type

'List Information
Private Type List
    ListID As String * 4
    dwSize As Long
    ListType As String * 4
End Type

'Stream StreamHeader information
Private Type AVIStreamHeader
    HeaderID As String * 4
    dwBytesIn As Long
    fccHandler As String * 4
    dwHandle As String * 4
    dwFlags As Long
    dwReserved1 As Long
    dwInitialFrames As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
End Type

'Stream Format
Private Type StreamFormat
    StreamFormatID As String * 4
    dwStreamSize As Long
End Type

Private Type LumpData
    LumpID As String * 4
    LumpSize As Long
End Type

Private Type FramesInfo
    ID As Integer
    FrameKey As String
End Type

Public Err_Code(7) As String        'Holds a array of error strings
Private Avi_FilePtr As Long         'AVI File Pointer
Private FilePos As Long             'Positon in the file
Public AviAttr As Integer           'Avi Attrib = 1 has Video Attrib = 2 has video and audio
Public sPallete As String           'Pallet Information
'Types
Private AviHead As AVIHeader
Public AVIInfo As AVIInfo
Private AviList As List
Public AviStream As AVIStreamHeader
Private StreamFormat As StreamFormat
Private ImgData As LumpData
Private TWaveInfo As WavHeadInfo
Private TWaveHead As WavHeader
Public TFramesInfo() As FramesInfo
Public Bmp As BITMAPINFOHEADER

Public Sub CloseAviFile()
    Close #Avi_FilePtr
    
    'Close the avi file and clean up garbage
    'Clear out the types we used
    ZeroMemory AviHead, Len(AviHead)
    ZeroMemory AVIInfo, Len(AVIInfo)
    ZeroMemory AviList, Len(AviList)
    ZeroMemory AviStream, Len(AviStream)
    ZeroMemory StreamFormat, Len(StreamFormat)
    ZeroMemory ImgData, Len(ImgData)
    ZeroMemory Bmp, Len(Bmp)
    
    sPallete = ""
    Erase TFramesInfo
    AviAttr = 0
    FilePos = 0
End Sub

Function OpenAviFile(lzFile As String) As Integer
Dim fp As Long
Dim BmpData As String
Dim sCount As Long
Dim p As Long

    Call CloseAviFile
    OpenAviFile = -1
    Avi_FilePtr = FreeFile
    
    'Main code for reading the information out of the AVI file
    'Note that at the moment all this code will do is read uncompressed AVI Files. ie 24- Bit
    
    Open lzFile For Binary As #Avi_FilePtr
        'Get AVI Fileheader info
        Get #Avi_FilePtr, , AviHead
        
        'Check for vaild RIFF and AVI ID
        'Get RIFF Tag all AVI will need to start with this
        If (AviHead.dwType <> "RIFF") Then
            OpenAviFile = 0
            Exit Function
        End If
        
        'All AVI files require the AVI Tag
        If Trim(AviHead.dwTAG) <> "AVI" Then
            OpenAviFile = 1
            Exit Function
        End If
        
        FilePos = Seek(Avi_FilePtr) + Len(AviHead)
        'Get Avi Information
        Get #Avi_FilePtr, FilePos, AVIInfo
        'Seek #Avi_FilePtr, (FilePos + AviInfo.Reserved) + 8
        FilePos = FilePos + 64
        Seek #Avi_FilePtr, FilePos
        
        'Loop tho the AVI Lists
        Do
            Get #Avi_FilePtr, , AviList
            
            If (AviList.ListType = "strl") Then
                FilePos = Loc(Avi_FilePtr) + 1
                'Get AVI Stream
                Get #Avi_FilePtr, , AviStream
               
                'Check for Stream ID
                If (AviStream.HeaderID <> "strh") Then
                    OpenAviFile = 2
                    Exit Function
                End If
            End If
            
            'We found a Video Stream
            If (AviStream.fccHandler = "vids") Then
                Seek #Avi_FilePtr, FilePos + AviStream.dwBytesIn + 8
                Get #Avi_FilePtr, , StreamFormat
                'Check for vaild StreamForm
                If (StreamFormat.StreamFormatID <> "strf") Then
                    OpenAviFile = 3
                    Exit Function
                Else
                    'Position ours selfs at the start of the video stream
                    FilePos = Seek(Avi_FilePtr)
                End If
                
                'Get Bitmap Info
                Get #Avi_FilePtr, , Bmp
                
                'Check if this an uncompressed AVI file
                If (Bmp.biCompression <> 0) Then OpenAviFile = 7: Exit Function
                
                'Check for vaild Bitmap color depths
                Select Case Bmp.biBitCount
                    Case 16, 24, 32
                    Case 8
                        sPallete = Space(1024)
                        Get #Avi_FilePtr, , sPallete
                    Case 4
                        sPallete = Space(64)
                        Get #Avi_FilePtr, , sPallete
                    Case Else
                        OpenAviFile = 4
                        Exit Function
                End Select
                
                Seek #Avi_FilePtr, FilePos + StreamFormat.dwStreamSize
                AviAttr = 1 'We have video stream
                
            ElseIf (AviStream.fccHandler = "auds") Then
                'Audio Stream Found I not got this working yet but I try and fix it for next time
                Seek #Avi_FilePtr, FilePos + 64
                Get #Avi_FilePtr, , TWaveHead
                Seek #Avi_FilePtr, (FilePos + AviList.dwSize - 4)
                AviAttr = AviAttr + 1   'We have Audio stream
            Else
                'AVI only supports 2 streams Video and Audio so we must exit
                OpenAviFile = 5
                Exit Function
            End If
        
            sCount = sCount + 1
        Loop Until (sCount >= AVIInfo.dwStreams)

        Do
            Get #Avi_FilePtr, , AviList
            Select Case UCase(AviList.ListID)
                'Skip over any junk
                Case "JUNK", "VEDT"
                    FilePos = Loc(Avi_FilePtr) + AviList.dwSize - 3
                    Seek #Avi_FilePtr, FilePos
                'If we find anything else just exit as this version does not support extras
                Case "LIST"
                    Exit Do
                Case Else
                    Exit Do
            End Select
        Loop
        
        'Loop tho all the frames and ident each one, this can be video, audio or compressed data
        sCount = 0
        Do Until (sCount = AVIInfo.dwTotalFrames)
            ReDim Preserve TFramesInfo(sCount)
            Get #Avi_FilePtr, , ImgData
            Select Case UCase(Right(ImgData.LumpID, 2))
                'Uncompressed Video Data
                Case "DB"
                    TFramesInfo(sCount).ID = 1
                    TFramesInfo(sCount).FrameKey = ":" & Seek(Avi_FilePtr) & "," & ImgData.LumpSize
                    'Extract Uncompressed Video Data
                    BmpData = Space(ImgData.LumpSize)
                    Get #Avi_FilePtr, , BmpData
                    p = p + Len(BmpData)
                    BmpData = ""
                Case "WB"
                    'Audio we do not support audio but we keep still store it
                    TFramesInfo(sCount).ID = 2
                    TFramesInfo(sCount).FrameKey = ":" & Seek(Avi_FilePtr) & "," & ImgData.LumpSize
                Case Else
                
            End Select
           sCount = sCount + 1
        Loop
        
        'Cleanup
        sCount = 0
        
End Function

Public Function GetStreamData(sOffset As Long, dSize As Long) As String
Dim sData As String
    
    'This function Extracts the Frame Data
    'Seek to Data Offset
    Seek #Avi_FilePtr, sOffset
    'Resize sData Buffer so we can store it's data
    sData = Space(dSize)
    'Get The Data
    Get #Avi_FilePtr, , sData
    'Return the Data
    GetStreamData = sData
    sData = ""
End Function

Public Sub InitAvi()
    'This sub just inits the error codes and Closes the AVI files and cleans up
    Call InitErrorCodes
    Call CloseAviFile
End Sub

Public Sub InitErrorCodes()
    'Our AVI Error codes
    Err_Code(0) = "No valid RIFF ID was found."
    Err_Code(1) = "No valid AVI tag found."
    Err_Code(2) = "Error locating stream header."
    Err_Code(3) = "No valid Stream Form was found."
    Err_Code(4) = "Only 16, 24, 32 -Bit AVI's are Supported."
    Err_Code(5) = "Unknown Stream found or not supported."
    Err_Code(6) = "Audio streams are not supported in this version"
    Err_Code(7) = "Compressed AVI's are not supported"
    
End Sub
