VERSION 5.00
Begin VB.Form LevelLoader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load Level"
   ClientHeight    =   3735
   ClientLeft      =   2565
   ClientTop       =   1455
   ClientWidth     =   3135
   ControlBox      =   0   'False
   Icon            =   "LevelLoader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileSelect 
      Height          =   3015
      Left            =   120
      Pattern         =   "*.til"
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   5
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
End
Attribute VB_Name = "LevelLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    PlayScreen.Enabled = True
    PlayScreen.Show
End Sub

Private Sub FileSelect_DblClick()
    Dim FileNum As Integer
    FileNum = FreeFile
    Open App.Path & "\Levels\" & FileSelect.FileName For Random As #FileNum
        For c = 1 To 20
        For d = 1 To 10
            Get #FileNum, ((c - 1) * 20) + d, MemBoard.Tiles(c, d)
        Next
        Next
    Close #FileNum
    PlayScreen.Enabled = True
    Unload Me
    DoEvents
    NewGame
End Sub

Private Sub Form_Load()
    FileSelect.Path = App.Path & "\Levels"
    FileSelect.Pattern = "*.til"
    FileSelect.Refresh
End Sub
