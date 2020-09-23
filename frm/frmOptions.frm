VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   1485
      TabIndex        =   2
      Top             =   2160
      Width           =   1015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   315
      TabIndex        =   1
      Top             =   2160
      Width           =   1015
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "#"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   315
      Width           =   2160
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UnloadControls()
Dim x As Integer
    x = Opt1.Count - 1
    Do Until (x = 0)
        Unload Opt1(x)
        x = x - 1
    Loop
End Sub

Private Sub LoadControls(Count As Integer)
Dim x As Integer
    Call UnloadControls
    For x = 1 To Count
        Load Opt1(x)
        Opt1(x).Top = Opt1(x - 1).Top + Opt1(x).Height + 5
        Opt1(x).Visible = True
    Next
End Sub

Private Sub SetOptions()
    If (ExportType = 1) Then
        LoadControls 4
        frmOptions.Caption = "Export Bitmap"
        'Set the option captions captions
        Opt1(0).Caption = "16 Color Bitmap"
        Opt1(1).Caption = "255 Color Bitmap"
        Opt1(2).Caption = "16-Bit Bitmap"
        Opt1(3).Caption = "24-Bit Bitmap"
        Opt1(4).Caption = "32-Bit Bitmap"
        Opt1(3).value = True
    ElseIf (ExportType = 2) Then
        LoadControls 2
        frmOptions.Caption = "Export TGA"
        'Set the option captions captions
        Opt1(0).Caption = "16-Bit TGA"
        Opt1(1).Caption = "24-Bit TGA"
        Opt1(2).Caption = "32-Bit TGA"
        Opt1(1).value = True
    Else
        frmOptions.Caption = "Export Icon/Cursor"
        LoadControls 1
        Opt1(0).Caption = "Icon-Resource"
        Opt1(1).Caption = "Cursor-Resource"
        Opt1(0).value = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmOptions
End Sub

Private Sub cmdOK_Click()
    Button_Press = 1
    Call UnloadControls
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Button_Press = 0
    Call SetOptions
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub Opt1_Click(Index As Integer)
    ExportOption = Index
End Sub
