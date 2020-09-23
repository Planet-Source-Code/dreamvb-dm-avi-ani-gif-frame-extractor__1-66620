VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM AVI, ANI, GIF Extractor V1.2"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10305
      TabIndex        =   23
      Top             =   5850
      Width           =   10365
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "DM AVI, ANI, GIF Extractor V1.2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   24
         Top             =   60
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CDG 
      Left            =   9600
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8895
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":09F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5565
      Left            =   3390
      ScaleHeight     =   5505
      ScaleWidth      =   6825
      TabIndex        =   1
      Top             =   240
      Width           =   6885
      Begin VB.PictureBox pViewer 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   930
         Left            =   4470
         ScaleHeight     =   930
         ScaleWidth      =   870
         TabIndex        =   22
         Top             =   135
         Width           =   870
      End
      Begin VB.PictureBox pInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5430
         Left            =   30
         ScaleHeight     =   5430
         ScaleWidth      =   6750
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   6750
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AVI Information::"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   945
            TabIndex        =   21
            Top             =   285
            Width           =   2325
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   60
            Picture         =   "frmmain.frx":0D48
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   2010
            TabIndex        =   20
            Top             =   4020
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   2010
            TabIndex        =   19
            Top             =   3735
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   780
            TabIndex        =   18
            Top             =   4020
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   780
            TabIndex        =   17
            Top             =   3735
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   2010
            TabIndex        =   16
            Top             =   3390
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   780
            TabIndex        =   15
            Top             =   3390
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   2010
            TabIndex        =   14
            Top             =   3105
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   780
            TabIndex        =   13
            Top             =   3105
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   2010
            TabIndex        =   12
            Top             =   2835
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   780
            TabIndex        =   11
            Top             =   2805
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   2010
            TabIndex        =   10
            Top             =   2505
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   780
            TabIndex        =   9
            Top             =   2505
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   2010
            TabIndex        =   8
            Top             =   2040
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   780
            TabIndex        =   7
            Top             =   2040
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2010
            TabIndex        =   6
            Top             =   1770
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   780
            TabIndex        =   5
            Top             =   1770
            Width           =   90
         End
         Begin VB.Label lblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   780
            TabIndex        =   4
            Top             =   1260
            Width           =   90
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   780
            TabIndex        =   3
            Top             =   990
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   3690
      Left            =   60
      TabIndex        =   0
      Top             =   225
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   6509
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Resource"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuexport 
      Caption         =   "&Export"
      Begin VB.Menu mnuBitmap 
         Caption         =   "&Bitmap"
         Begin VB.Menu mnusel 
            Caption         =   "Selected"
         End
         Begin VB.Menu mnuAll 
            Caption         =   "&All"
         End
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 1.1
'Now supportd
'16 Color, 8 bit color
'16,32 Bit true color
'Export selected Frames to TGA

'Version 1.2
'Support to extract frames form animated cursors
'Exports animated cursors frames as icons or cursors
'GIF Support, export GIF Frames as GIF, Bitmap, TGA
'Added Gif Frame info for each frame in the GIF File

Option Explicit
Private Res_Filename As String
Private Resource_Type As Integer

Private Sub GifReource()
Dim iRet As Integer
Dim TvNode As Node

    'Open AVI Resource
    Call Init
    GifFilename = Res_Filename
    iRet = OpenGIF()
    mnuexport.Enabled = iRet
    
    If (iRet <> -1) Then
        'Looks like an error was found.
        MsgBox GifAbort(iRet), vbInformation, "Error_" & iRet
        'Close AVI file
        Call CleanGlobal
    Else
        'Fill the treeview
        Call InitTvNodes(3)
        Set TvNode = tv1.Nodes(1)
        TvNode.Selected = True
        tv1_NodeClick TvNode
        Set TvNode = tv1.Nodes(2)
        TvNode.Selected = True
        Set TvNode = Nothing
        mnuBitmap.Caption = "Bitmap"
    End If
End Sub

Private Sub AviReource()
Dim iRet As Integer
Dim TvNode As Node

    'Open AVI Resource
    Call InitAvi
    iRet = OpenAviFile(Res_Filename)
    mnuexport.Enabled = iRet
    
    If (iRet <> -1) Then
        'Looks like an error was found.
        MsgBox Err_Code(iRet), vbInformation, "Error_" & iRet
        'Close AVI file
        Call CloseAviFile
    Else
        'Fill the treeview
        Call InitTvNodes(1)
        Set TvNode = tv1.Nodes(1)
        TvNode.Selected = True
        tv1_NodeClick TvNode
        Set TvNode = tv1.Nodes(2)
        TvNode.Selected = True
        Set TvNode = Nothing
        mnuBitmap.Caption = "Bitmap"
    End If
End Sub

Private Sub AniReource()
Dim iRet As Integer
Dim TvNode As Node

    'Open ANI Resource
    iRet = OpenAniFile(Res_Filename)
    mnuexport.Enabled = iRet
    
    If (iRet <> -1) Then
        'Looks like an error was found.
        MsgBox AniErr(iRet), vbInformation, "Error_" & iRet
        Call CleanUp
    Else
        'Fill the treeview
        Call InitTvNodes(2)
        Set TvNode = tv1.Nodes(1)
        TvNode.Selected = True
        tv1_NodeClick TvNode
        Set TvNode = tv1.Nodes(2)
        TvNode.Selected = True
        Set TvNode = Nothing
        mnuBitmap.Caption = "Icon"
    End If
    
End Sub

Private Sub CenterViewer()
    pViewer.Left = (pBase.ScaleWidth - pViewer.Width) \ 2
    pViewer.Top = (pBase.ScaleHeight - pViewer.Height) \ 2
End Sub


Private Sub InitTvNodes(rType As Integer)
Dim x As Long
    With tv1
        .Indentation = 60
        .Nodes.Clear
        .Nodes.Add , tvwFirst, "TOP", GetFilename(Res_Filename), 1, 1
        
        'AVI Resource
        If (rType = 1) Then
            .Nodes.Add 1, tvwChild, "STR", "STREAMS", 2, 2
            .Nodes.Add 2, tvwChild, "VID", "Video", 2, 2
            .Nodes.Add 2, tvwChild, "AUD", "Audio", 2, 2
            
            For x = 0 To UBound(TFramesInfo)
                If TFramesInfo(x).ID = 1 Then
                    .Nodes.Add 3, tvwChild, TFramesInfo(x).FrameKey, "Frame_" & x + 1, 3, 3
                ElseIf TFramesInfo(x).ID = 2 Then
                    .Nodes.Add 4, tvwChild, TFramesInfo(x).FrameKey, "Frame_" & x + 1, 4, 4
                End If
            Next x
        End If
        
        'ANI Resource
        If (rType = 2) Then
            .Nodes.Add 1, tvwChild, "ANI", "Frames", 2, 2
            'Add the frames
            For x = 0 To AniInfo.cFrames - 1
                .Nodes.Add 2, tvwChild, ":" & x, "Frame_" & x + 1, 3, 3
            Next x
        End If
                
        'Gif Files
        If (rType = 3) Then
            .Nodes.Add 1, tvwChild, "GIF", "Frames", 2, 2
            For x = 0 To TGifHeadInfo.FrameCount - 1
                .Nodes.Add 2, tvwChild, ":" & x, "Frame_" & x + 1, 3, 3
            Next x
        End If
    End With
    
    x = 0
    
End Sub

Private Sub ShowResInfo(rType As Integer)
    'This sub just shows some basic information for the resource type, AVI, ANI and GIF
    
    lblA(0).Caption = Res_Filename
    
    Select Case rType
        Case 1 'AVI
            Label1.Caption = "AVI Information:"
            lblInfo(1).Caption = "Width:"
            lblInfo(2).Caption = "Height:"
            lblInfo(3).Caption = "Streams:"
            lblInfo(4).Caption = "Frames:"
            lblInfo(5).Caption = "Rate:"
            lblInfo(6).Caption = "Sample Size:"
            lblInfo(7).Caption = "Has Sound:"
            lblInfo(8).Caption = "Has Video:"
            lblA(1).Caption = AVIInfo.dwWidth
            lblA(2).Caption = AVIInfo.dwHeight
            lblA(3).Caption = AVIInfo.dwStreams
            lblA(4).Caption = AVIInfo.dwTotalFrames
            lblA(5).Caption = AviStream.dwRate & " frames/second"
            lblA(6).Caption = Bmp.biBitCount & " Bit"
            lblA(7).Caption = (AviAttr >= 2)
            lblA(8).Caption = (AviAttr >= 1)
        Case 2 'ANI
            Label1.Caption = "ANI Information:"
            lblInfo(1).Caption = "Width:"
            lblInfo(2).Caption = "Height:"
            lblInfo(3).Caption = "Frames:"
            lblInfo(4).Caption = "Pixel Count:"
            lblInfo(5).Caption = "Title:"
            lblInfo(6).Caption = "Credits:"
            lblInfo(7).Caption = "Filesize:"
            lblInfo(8).Caption = "Color Format:"
            lblA(1).Caption = AniInfo.wWidth
            lblA(2).Caption = AniInfo.wHeight
            lblA(3).Caption = AniInfo.cFrames
            lblA(4).Caption = AniInfo.PixelCount
            lblA(5).Caption = AniInfo.Title
            lblA(6).Caption = AniInfo.Credits
            lblA(7).Caption = AniInfo.Filesize
            lblA(8).Caption = AniInfo.StrFormat
        Case 3 'GIF
            Label1.Caption = "GIF Information:"
            lblInfo(1).Caption = "Version:"
            lblInfo(2).Caption = "FrameCount:"
            lblInfo(3).Caption = "Width:"
            lblInfo(4).Caption = "Height:"
            lblInfo(5).Caption = "Colors:"
            lblInfo(6).Caption = "Repeat:"
            lblInfo(7).Caption = "Appliaction:"
            lblInfo(8).Caption = "Comment:"
            lblA(1).Caption = TGifHeadInfo.Version
            lblA(2).Caption = TGifHeadInfo.FrameCount
            lblA(3).Caption = TGifHeadInfo.Width
            lblA(4).Caption = TGifHeadInfo.Height
            lblA(5).Caption = TGifHeadInfo.dColors
            lblA(6).Caption = TGifHeadInfo.dRepeat
            lblA(7).Caption = TGifHeadInfo.sAppName
            lblA(8).Caption = TGifHeadInfo.sComment
    End Select
    
End Sub

Private Sub Form_Resize()
    tv1.Height = (frmmain.ScaleHeight - pBar.Height - tv1.Top) - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Public Sub DisplayImage(cTvKey As String, iPicBox As PictureBox)
Dim dOffset As Long
Dim dSize As Long
Dim vData As Variant
Dim TmpSaveName As String
Dim f As Long

    'Temp Filename
    TmpSaveName = FixPath(App.Path) & "tmpimg.bmp"
    'Get Data Information
    vData = Split(cTvKey, ",")  'Store cTvKey
    dOffset = CLng(vData(0))    'Data Offset
    dSize = CLng(vData(1))      'Data Size
    
    If (Bmp.biBitCount = 8) Or (Bmp.biBitCount = 4) Then
        Call Save8BitBmp(TmpSaveName, GetStreamData(dOffset, dSize), AVIInfo.dwWidth, AVIInfo.dwHeight, Bmp.biBitCount, sPallete)
    Else
        Call SaveBmp24Lump(TmpSaveName, GetStreamData(dOffset, dSize), AVIInfo.dwWidth, AVIInfo.dwHeight, Bmp.biBitCount)
    End If
    
    'Display the temp image in the Picturebox
    iPicBox.Picture = LoadPicture(TmpSaveName)
    Kill TmpSaveName

    'Clean up Garbage
    dOffset = 0
    dSize = 0
    TmpSaveName = ""
    Erase vData
End Sub

Sub DisplayIcon(Index As String, iPicBox As PictureBox)
Dim TmpName As String
    TmpName = FixPath(App.Path) & "tmp.ico"
    If Not ExtractAniFrame(Res_Filename, CInt(Index), TmpName) Then Exit Sub
    pViewer.Picture = LoadPicture(TmpName)
    Kill TmpName
    'Clean Var
    TmpName = ""
End Sub

Sub DisplayGif(Index As String, iPicBox As PictureBox)
Dim TmpName As String
Dim sPos As Long, ePos As Long
    'Display some frame info
    pBase.Print ""
    pBase.Print " Erase Method:", EraseMethod(FrameInfo(Index).gEreaseMethod)
    pBase.Print " Time Delay:", , FrameInfo(Index).gDelayTime
    pBase.Print " Top:", , FrameInfo(Index).gTop
    pBase.Print " Left:", , FrameInfo(Index).gLeft
    pBase.Print " Trans Color", , FrameInfo(Index).TransColor
    
    TmpName = FixPath(App.Path) & "tmp.gif"
    'SaveSingleFrame
    sPos = FrameInfo(Index).ImgStartPos
    ePos = FrameInfo(Index).ImgEndPos
    Call SaveSingleFrame(TmpName, sPos, ePos)
    pViewer.Picture = LoadPicture(TmpName)
    Kill TmpName
    'Clean vars
    sPos = 0: ePos = 0: TmpName = ""
End Sub

Private Sub mnuabout_Click()
    MsgBox lblTitle.Caption & vbCrLf & "Created By DreamVb" & vbCrLf & vbTab & "Please Vote.", vbInformation, "About"
End Sub

Private Sub mnuAll_Click()
Dim FolName As String
Dim x As Integer
Dim StrTmp As String
Dim vData As Variant
Dim dOffset As Long
Dim dSize As Long
Dim idx As Long
Dim lzPath As String, lzFile As String
Dim TmpExt As String

    FolName = GetFolder(frmmain.hWnd, "Export")
    If Len(FolName) = 0 Then Exit Sub
    lzPath = FixPath(FolName)
    
    'Save all Bitmaps for AVI Resource type
    If (Resource_Type = 1) Then
        'Export all the images
        For x = 1 To tv1.SelectedItem.Children
            'Node Key
            StrTmp = Mid(tv1.Nodes(tv1.SelectedItem.Index + 1 + x).Key, 2)
            'Get the Data Info
            vData = Split(StrTmp, ",")
            dOffset = CLng(vData(0))    'Data Offset
            dSize = CLng(vData(1))      'Data Size
            'Save Bitmap
            lzFile = lzPath & tv1.Nodes(tv1.SelectedItem.Index + 1 + x).Text & ".bmp"
            Select Case Bmp.biBitCount
                Case 4, 8
                    Call Save8BitBmp(lzFile, GetStreamData(dOffset, dSize), AVIInfo.dwWidth, AVIInfo.dwHeight, Bmp.biBitCount, sPallete)
                Case 16, 24, 32
                    Call SaveBmp24Lump(lzFile, GetStreamData(dOffset, dSize), AVIInfo.dwWidth, AVIInfo.dwHeight, Bmp.biBitCount)
            End Select
        Next x
        
        MsgBox x - 1 & " frames have been successfully exported to:" & vbCrLf & FixPath(FolName), vbInformation, frmmain.Caption
        'Clean up Garabge
        Erase vData
        dOffset = 0
        dSize = 0
    End If
    
    'Extract all icons or cursors
    If (Resource_Type = 2) Then
        ExportType = 3
        frmOptions.Show vbModal, frmmain
        If (Button_Press = 0) Then Exit Sub
        If (ExportOption + 1) = 1 Then TmpExt = ".ico" Else TmpExt = ".cur"
        'Save the resource
        For x = 1 To tv1.SelectedItem.Children
            StrTmp = Mid(tv1.Nodes(tv1.SelectedItem.Index + x).Key, 2)
            lzFile = tv1.Nodes(tv1.SelectedItem.Index + x).Text & TmpExt
            Call ExtractAniFrame(Res_Filename, CInt(StrTmp), lzPath & lzFile, ExportOption + 1)
        Next x
        
        If (ExportOption + 1) = 1 Then StrTmp = "Icons" Else StrTmp = "Cursors"
        MsgBox x - 1 & " " & StrTmp & " have been successfully exported to:" & vbCrLf & FixPath(FolName), vbInformation, frmmain.Caption
    End If
    
    'Extract all Gif frames
    If (Resource_Type = 3) Then
        For x = 1 To tv1.SelectedItem.Children
            StrTmp = Mid(tv1.Nodes(tv1.SelectedItem.Index + x).Key, 2)
            lzFile = tv1.Nodes(tv1.SelectedItem.Index + x).Text & ".gif"
            dOffset = FrameInfo(StrTmp).ImgStartPos
            dSize = FrameInfo(StrTmp).ImgEndPos
            Call SaveSingleFrame(lzPath & lzFile, dOffset, dSize)
        Next x
        
        MsgBox x - 1 & " Gif files have been successfully exported to:" & vbCrLf & FixPath(FolName), vbInformation, frmmain.Caption
    End If
    'Clean up Garabge
    lzFile = ""
    StrTmp = ""
    FolName = ""
    TmpExt = ""
    x = 0
End Sub

Private Sub mnuexit_Click()
    Call CloseAviFile
    tv1.Nodes.Clear
    Unload frmmain
End Sub

Private Sub mnuOpen_Click()
On Error GoTo CanErr:

    With CDG
        .CancelError = True
        .DialogTitle = "Open Resource"
        .Filter = "AVI Files (*.avi)|*.avi|Animated Cursor (*.ani)|*.ani|GIF Files (*.gif)|*.gif|"
        .InitDir = App.Path
        
        .ShowOpen
        
        Resource_Type = 0
         Res_Filename = .Filename
        .Filename = ""
        'Hide the information display
        pInfo.Visible = False
        Resource_Type = .FilterIndex
        
        Select Case .FilterIndex
            Case 1  'AVI
                Call AviReource
            Case 2  'Animated Cursor
                Call AniReource
            Case 3  'Animated GIF
                Call GifReource
            Case Else
                MsgBox "Resource not supported."
        End Select
    End With
    
    Exit Sub
CanErr:
    If Err = cdlCancel Then
        Err.Clear
    End If
End Sub

Private Sub mnusel_Click()
On Error GoTo CanErr:
Dim idx As Integer
Dim sPos As Long, ePos As Long

    With CDG
        .CancelError = True
        .DialogTitle = "Save As"
        
        If (Resource_Type = 1) Then
            .Filter = "Bitmap Files (*.bmp)|*.bmp|TGA Files (*.tga)|*.tga|"
        End If
        
        If (Resource_Type = 2) Then
            .Filter = "Bitmap Files (*.bmp)|*.bmp|TGA Files (*.tga)|*.tga|Icon Resource (*.ico)|*.ico|Cursor Resource (*.cur)|*.cur|"
        End If
        
        If (Resource_Type = 3) Then
            .Filter = "Bitmap Files (*.bmp)|*.bmp|TGA Files (*.tga)|*.tga|GIF Files(*.gif)|*.gif|"
        End If
        

        .Filename = tv1.SelectedItem.Text
        .ShowSave
        ExportType = .FilterIndex
        
        If (ExportType = 1) Or (ExportType = 2) Then
            frmOptions.Show vbModal, frmmain
            'No choice was made so we do nothing
            If (Button_Press = 0) Then Exit Sub
        End If

        Select Case ExportType
            Case 1
                'Save Bitmap
                Call SaveBmp(pViewer, .Filename, ExportOption)
            Case 2
                'Save TGA
                Call SaveTGA(pViewer, .Filename, ExportOption)
            Case 3, 4
                'Get the index of the treeview for the frame resource
                idx = Right(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1)
                'Save Icon or Cursor
                If (Resource_Type = 2) Then Call ExtractAniFrame(Res_Filename, idx, .Filename, (ExportType - 3) + 1)
                
                'Save GIF File
                If (Resource_Type = 3) Then
                    sPos = FrameInfo(idx).ImgStartPos
                    ePos = FrameInfo(idx).ImgEndPos
                    Call SaveSingleFrame(.Filename, sPos, ePos)
                End If
                
        End Select
        
        sPos = 0
        ePos = 0
        .Filename = ""
    End With
    
    Exit Sub
CanErr:
    If Err = cdlCancel Then
        Err.Clear
    End If
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim cKey As String, pKey As String
On Error Resume Next
    
    If (tv1.Nodes.Count = 0) Then Exit Sub
    cKey = tv1.SelectedItem.Key
    
    pBase.Cls
    
    pInfo.Visible = (cKey = "TOP")
    
    If Left(cKey, 1) <> ":" Then
        pViewer.Visible = False
        mnusel.Enabled = False
        mnuAll.Enabled = False
        
        Select Case cKey
            Case "TOP"
                'Show Resource Information
                Call ShowResInfo(Resource_Type)
            Case "VID", "ANI", "GIF"
                mnuAll.Enabled = True
            Case "AUD"
        End Select
    Else
        pKey = tv1.SelectedItem.Parent.Key
        
        If (pKey = "VID") Then
            mnusel.Enabled = True
            mnuAll.Enabled = False
            pViewer.Visible = True
            'Strip way colon :
            cKey = Right(cKey, Len(cKey) - 1)
            'Display the Frame
            Call DisplayImage(cKey, pViewer)
            Call CenterViewer
        ElseIf (pKey = "AUD") Then
            'Audio is still not supported
            mnusel.Enabled = False
            mnuAll.Enabled = False
           cKey = Right(cKey, Len(cKey) - 1)
            Exit Sub
        ElseIf (pKey = "ANI") Then
            mnusel.Enabled = True
            mnuAll.Enabled = False
            pViewer.Visible = True
            'Strip way colon :
            cKey = Right(cKey, Len(cKey) - 1)
            'Display the icon
            Call DisplayIcon(cKey, pViewer)
            Call CenterViewer
        ElseIf (pKey = "GIF") Then
            mnusel.Enabled = True
            mnuAll.Enabled = False
            pViewer.Visible = True
            cKey = Right(cKey, Len(cKey) - 1)
            Call DisplayGif(cKey, pViewer)
            Call CenterViewer
        End If
    End If
    
End Sub
