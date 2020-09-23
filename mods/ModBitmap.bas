Attribute VB_Name = "ImageExport"
Option Explicit

Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, _
    ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, _
    ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'TGA Information Header
Public Type TgaHeader
    IdentSize As Byte
    ColorType As Byte
    ImageType As Byte
    ColourMapStart As Integer
    ColourMapLength As Integer
    ColourMapBits As Byte
    value  As Integer
    yvalue As Integer
    Width As Integer
    Height As Integer
    BitCount  As Byte
    Descriptor As Byte
End Type

'Bitmap Information Header
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Bitmap File Header
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

'Bitmap Formats consts
Public Enum ColorDepth
    Bmp16 = 0
    Bmp255 = 1
    Bmp16Bit = 2
    Bmp24Bit = 3
    Bmp32Bit = 4
End Enum

Public Sub SaveBmp(iPicBox As PictureBox, Filename As String, Optional iColorFormat As Byte = 3)
Dim BmpFile As BITMAPFILEHEADER
Dim BmpInfo As BITMAPINFO
Dim PalSize As Long, Palette() As Long
Dim ClrValue As Integer, BitCount As Integer
Dim fp As Long
Dim x As Integer
Dim BmpW As Long, BmpH As Long

Dim ImgSize As Long
Dim ImageBits() As Byte
Dim iRet As Long
    
    BmpW = (iPicBox.ScaleWidth \ Screen.TwipsPerPixelX)
    BmpH = (iPicBox.ScaleHeight \ Screen.TwipsPerPixelY)
    
    'Used to ident the Color depth format to use
    Select Case iColorFormat
        Case Bmp16
            '4 Bitmap
            PalSize = 15
            ClrValue = 16
            BitCount = 4
        Case Bmp255
            '255 Bitmap
            PalSize = 255
            ClrValue = 256
            BitCount = 8
        Case Bmp16Bit
            '16Bit true color
            PalSize = 0
            BitCount = 16
        Case Bmp24Bit
            '24Bit true color
            PalSize = 0
            BitCount = 24
        Case Bmp32Bit
            '32Bit true color
            PalSize = 32
            BitCount = 32
    End Select
    
    'Fill in Bitmap FileInfo
    With BmpInfo
        .bmiHeader.biBitCount = BitCount
        .bmiHeader.biClrImportant = ClrValue
        .bmiHeader.biClrUsed = ClrValue
        .bmiHeader.biCompression = 0
        .bmiHeader.biHeight = BmpH
        .bmiHeader.biPlanes = 1
        .bmiHeader.biSize = Len(.bmiHeader)
        .bmiHeader.biWidth = BmpW
    End With
    
    With BmpInfo.bmiHeader
        'Get the bitmap bits using GetDIBits
        iRet = GetDIBits(iPicBox.hDC, iPicBox.Image, 0, .biHeight, ByVal 0, BmpInfo, 0)
        'Store the image size to hold the Palette
        ImgSize = (.biSizeImage / .biHeight)
        'Resize image bits to hold the data
        ReDim ImageBits(1 To ImgSize, 1 To .biHeight)
        iRet = GetDIBits(iPicBox.hDC, iPicBox.Image, 0, .biHeight, ImageBits(1, 1), BmpInfo, 0)
        'Rezize pallete table
        ReDim Palette(PalSize)
        'Build Pallete table
        For x = 0 To PalSize
            Palette(x) = BmpInfo.bmiColors(x)
        Next x
    End With
    
    'Fill Bitmap File Header Info
    With BmpFile
        .bfType = &H4D42
        .bfOffBits = &H36 + (PalSize + 1) * &H4
        .bfSize = .bfOffBits
    End With
    
    'Get a free filename
    fp = FreeFile
    'Write the Bitmap file
    Open Filename For Binary As #fp
        'Bitmap File header info
        Put #fp, , BmpFile
        'Bitmap image info
        Put #fp, , BmpInfo.bmiHeader
        'Bitmaps Palette
        Put #fp, , Palette()
        'Bitmap data it's self
        Put #fp, , ImageBits()
    Close #fp
    
    'Clear up
    Erase Palette
    Erase BmpInfo.bmiColors
    
    ZeroMemory BmpInfo.bmiHeader, Len(BmpInfo.bmiHeader)
    ZeroMemory BmpFile, Len(BmpFile)
End Sub

Sub SaveBmp24Lump(Filename As String, BmpData As String, BmpW As Long, BmpH As Long, PixelBitCount As Integer)
Dim fp As Long
Dim BmpFile As BITMAPFILEHEADER
Dim BmpInfo As BITMAPINFOHEADER
Dim Palette As Long

    'Fill in bitmap file header info
    With BmpFile
        .bfType = &H4D42    'BM
        .bfOffBits = &H36 + 1 * &H4
        .bfSize = .bfOffBits
    End With
    
    'Fill in bitmap information
    With BmpInfo
        .biBitCount = PixelBitCount
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = 0
        .biHeight = BmpH
        .biPlanes = 1
        .biSize = 40
        .biWidth = BmpW
    End With
    
    fp = FreeFile
    
    Open Filename For Binary As #fp
        Put #fp, , BmpFile  'Bitmap File Header
        Put #fp, , BmpInfo  'Bitmap Information Header
        Put #fp, , Palette  'No Pallet required for 24-bit bitmaps
        Put #fp, , BmpData  'Bitmap Data
    Close #fp
    
    ZeroMemory BmpInfo, Len(BmpInfo)
    ZeroMemory BmpFile, Len(BmpFile)
    
End Sub

Sub Save8BitBmp(Filename As String, BmpData As String, BmpW As Long, BmpH As Long, PixelBitCount As Integer, sPal As String)
Dim fp As Long
Dim BmpFile As BITMAPFILEHEADER
Dim BmpInfo As BITMAPINFOHEADER
Dim PalSize As Integer
Dim cVal As Integer

    If (PixelBitCount = 8) Then PalSize = 255: cVal = 256
    If (PixelBitCount = 4) Then PalSize = 15: cVal = 16
    
    'Fill in bitmap file header info
    With BmpFile
        .bfType = &H4D42    'BM
        .bfOffBits = &H36 + (PalSize + 1) * &H4
        .bfSize = .bfOffBits
    End With
    
    'Fill in bitmap information
    With BmpInfo
        .biBitCount = PixelBitCount
        .biClrImportant = cVal
        .biClrUsed = cVal
        .biCompression = 0
        .biHeight = BmpH
        .biPlanes = 1
        .biSize = 40
        .biWidth = BmpW
    End With
    
    fp = FreeFile
    
    Open Filename For Binary As #fp
        Put #fp, , BmpFile  'Bitmap File Header
        Put #fp, , BmpInfo  'Bitmap Information Header
        Put #fp, , sPal     'Bitmaps Pallete information
        Put #fp, , BmpData  'Bitmap Data
    Close #fp
    
    ZeroMemory BmpInfo, Len(BmpInfo)
    ZeroMemory BmpFile, Len(BmpFile)
    
    cVal = 0
    PalSize = 0
End Sub

Public Sub SaveTGA(iPicBox As PictureBox, Filename As String, Optional iColorFormat As Byte = 1)
Dim TgaHead As TgaHeader
Dim BmpInfo As BITMAPINFO
Dim BitCount As Byte
Dim fp As Long
Dim BmpW As Long, BmpH As Long

Dim ImgSize As Long
Dim ImageBits() As Byte

    BmpW = (iPicBox.ScaleWidth \ Screen.TwipsPerPixelX)
    BmpH = (iPicBox.ScaleHeight \ Screen.TwipsPerPixelY)
    
    'Used to ident the Color depth format to use
    If (iColorFormat = 0) Then BitCount = 16
    If (iColorFormat = 1) Then BitCount = 24
    If (iColorFormat = 2) Then BitCount = 32

    'Fill in Bitmap FileInfo
    With BmpInfo.bmiHeader
        .biBitCount = BitCount
        .biCompression = 0
        .biHeight = BmpH
        .biPlanes = 1
        .biSize = 40
        .biWidth = BmpW
    End With
    
    With BmpInfo
        'Get the bitmap bits using GetDIBits
        GetDIBits iPicBox.hDC, iPicBox.Image, 0, .bmiHeader.biHeight, ByVal 0, BmpInfo, 0
        'Store the image size to hold the Palette
        ImgSize = (.bmiHeader.biSizeImage / .bmiHeader.biHeight)
        'Resize image bits to hold the data
        ReDim ImageBits(1 To ImgSize, 1 To .bmiHeader.biHeight)
        GetDIBits iPicBox.hDC, iPicBox.Image, 0, .bmiHeader.biHeight, ImageBits(1, 1), BmpInfo, 0
    End With
    
    'Fill TgaHead
    With TgaHead
        .BitCount = BitCount
        .ImageType = 2  'RGB
        .Height = BmpH
        .Width = BmpW
    End With
    
    'Get a free filename
    fp = FreeFile
    
    'Write the TGA file
    Open Filename For Binary As #fp
        Put #fp, , TgaHead
        Put #fp, , ImageBits()
    Close #fp
    
    'Clear up
    ZeroMemory BmpInfo, Len(BmpInfo)
    ZeroMemory TgaHead, Len(TgaHead)
End Sub

