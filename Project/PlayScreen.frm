VERSION 5.00
Begin VB.Form PlayScreen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Naughtz"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   DrawWidth       =   3
   Icon            =   "PlayScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PlayScreen.frx":08CA
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer UFOMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   -7
      TabIndex        =   0
      Top             =   5040
      Width           =   9015
      Begin VB.Image PlayerLogo 
         Height          =   750
         Index           =   0
         Left            =   2040
         Picture         =   "PlayScreen.frx":B058C
         Top             =   120
         Width           =   2250
      End
      Begin VB.Image PlayerLogo 
         Height          =   750
         Index           =   1
         Left            =   4920
         Picture         =   "PlayScreen.frx":B5E16
         Top             =   120
         Visible         =   0   'False
         Width           =   2250
      End
   End
   Begin VB.Image GreenUFO 
      Height          =   300
      Left            =   120
      Picture         =   "PlayScreen.frx":BB6A0
      Top             =   120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image RedUFO 
      Height          =   300
      Left            =   120
      Picture         =   "PlayScreen.frx":BBA7D
      Top             =   480
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image BlueUFO 
      Height          =   300
      Left            =   120
      Picture         =   "PlayScreen.frx":BBE5A
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu Menu_Game 
      Caption         =   "Game"
      Begin VB.Menu Menu_Game_New 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu_Game_Names 
         Caption         =   "Set Player Names"
      End
      Begin VB.Menu Menu_Game_Switch 
         Caption         =   "Switch Player"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu_Game_Preset 
         Caption         =   "Load Preset"
         Begin VB.Menu Menu_Game_Preset_Normal 
            Caption         =   "Normal"
         End
         Begin VB.Menu Menu_Game_Preset_Full 
            Caption         =   "Full"
         End
      End
      Begin VB.Menu Menu_Game_Load 
         Caption         =   "Load Level"
         Shortcut        =   ^L
      End
      Begin VB.Menu Menu_Game_Edit 
         Caption         =   "Level Editor"
         Shortcut        =   ^D
      End
      Begin VB.Menu Menu_Game_Options 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "Editor"
      Begin VB.Menu Menu_Edit_Tiles 
         Caption         =   "Tiles"
         Begin VB.Menu Menu_Edit_Tiles_Blank 
            Caption         =   "Blank"
            Checked         =   -1  'True
            Shortcut        =   ^B
         End
         Begin VB.Menu Menu_Edit_Tiles_None 
            Caption         =   "No Tile"
            Shortcut        =   ^E
         End
         Begin VB.Menu Menu_Edit_Tiles_Naught 
            Caption         =   "Naught"
            Shortcut        =   ^O
         End
         Begin VB.Menu Menu_Edit_Tiles_Cross 
            Caption         =   "Cross"
            Shortcut        =   ^X
         End
      End
      Begin VB.Menu Menu_Edit_Fill 
         Caption         =   "Fill"
         Shortcut        =   ^F
      End
      Begin VB.Menu Menu_Edit_Clear 
         Caption         =   "Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu_Edit_Game 
         Caption         =   "Back To Game"
         Shortcut        =   ^G
      End
      Begin VB.Menu Menu_Edit_Save 
         Caption         =   "Save Level"
      End
      Begin VB.Menu Menu_Edit_Load 
         Caption         =   "Load Level"
      End
   End
   Begin VB.Menu Menu_About 
      Caption         =   "About"
   End
   Begin VB.Menu Menu_Quit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "PlayScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Menu_Edit.Enabled = False
    Cls
    ClearMemStandardBoard
    MemClearBoard
    WriteBoard
    Player1 = "Naughts"
    Player2 = "Crosses"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TileX As Integer
    Dim TileY As Integer
    If CanNotPlay = True Then Exit Sub
    If UFOInUse = True Then Exit Sub
    If (X < BoardLeft) Or (X > BoardLeft + 400) Or (Y < BoardTop) Or (Y > BoardTop + 200) Then Exit Sub
    X = X - BoardLeft
    Y = Y - BoardTop
    X = X / 20
    Y = Y / 20
    X = Fix(X)
    Y = Fix(Y)
    TileX = X + 1
    TileY = Y + 1
    If UserMode = 0 And (Board.Tiles(TileX, TileY).Tile = Normal Or Board.Tiles(TileX, TileY).Tile = Slow) Then
        If CheckPlayerGo(TileX, TileY, True) = True Then Exit Sub
        Randomize
        If AllowUFO = True And Rnd > UFOOccurance Then
            Dim RndNum As Single
            Randomize
            RndNum = Rnd
            If RndNum > GreenLimit And RndNum < BlueLimit Then UFOType = UFOBlue
            If RndNum > BlueLimit And RndNum < RedLimit Then UFOType = UFOGreen
            If RndNum > RedLimit Then UFOType = UFORed
            On Error Resume Next
            If UFOType = UFOGreen Then
                Do Until Board.Tiles(UFOAttackX, UFOAttackY).Tile = Normal
                    UFOAttackX = Rnd(1) * 20 + 1
                    UFOAttackY = Rnd(1) * 10 + 1
                Loop
            Else
                Do Until Board.Tiles(UFOAttackX, UFOAttackY).Tile <> Normal And Board.Tiles(UFOAttackX, UFOAttackY).Tile <> Missing
                    UFOAttackX = Rnd(1) * 20 + 1
                    UFOAttackY = Rnd(1) * 10 + 1
                Loop
            End If
            On Error GoTo 0
            UFOInUse = True
            If UFOType = UFOGreen Then
                GreenUFO.Top = ((UFOAttackY - 1) * 20) + BoardTop
                GreenUFO.Left = BoardLeft - 50
                GreenUFO.Visible = True
            End If
            If UFOType = UFOBlue Then
                BlueUFO.Top = ((UFOAttackY - 1) * 20) + BoardTop
                BlueUFO.Left = BoardLeft - 50
                BlueUFO.Visible = True
            End If
            If UFOType = UFORed Then
                RedUFO.Top = ((UFOAttackY - 1) * 20) + BoardTop
                RedUFO.Left = BoardLeft - 50
                RedUFO.Visible = True
            End If
            UFOMove.Enabled = True
        End If
    ElseIf UserMode = 1 And Menu_Edit_Tiles_Blank.Checked = True Then
        MemBoard.Tiles(TileX, TileY).Tile = Normal
        MemClearBoard
        Cls
        WriteBoard
    ElseIf UserMode = 1 And Menu_Edit_Tiles_None.Checked = True Then
        MemBoard.Tiles(TileX, TileY).Tile = Missing
        MemClearBoard
        Cls
        WriteBoard
    ElseIf UserMode = 1 And Menu_Edit_Tiles_Naught.Checked = True Then
        MemBoard.Tiles(TileX, TileY).Tile = Naught
        MemClearBoard
        Cls
        WriteBoard
    ElseIf UserMode = 1 And Menu_Edit_Tiles_Cross.Checked = True Then
        MemBoard.Tiles(TileX, TileY).Tile = Cross
        MemClearBoard
        Cls
        WriteBoard
    End If
    X = X * 20
    Y = Y * 20
    
End Sub

Private Sub Menu_Editor_Click()

End Sub

Private Sub Menu_About_Click()
PlayScreen.Enabled = False
About.Show
End Sub

Private Sub Menu_Edit_Clear_Click()
    Cls
    ClearMemBlankBoard
    MemClearBoard
    WriteBoard
End Sub

Private Sub Menu_Edit_Fill_Click()
    Cls
    ClearBoard
    MemBoard = Board
    MemClearBoard
    WriteBoard
End Sub

Private Sub Menu_Edit_Game_Click()
Menu_Edit.Enabled = False
Menu_Game.Enabled = True
Cls
MemClearBoard
WriteBoard
UserMode = 0
End Sub

Private Sub Menu_Edit_Load_Click()
PlayScreen.Enabled = False
LevelLoader.Show
End Sub

Private Sub Menu_Edit_Save_Click()
Dim FileNum As Integer
Dim FileName As String
FileName = InputBox("Please Enter The Name Of The Level")
FileNum = FreeFile
If FileName = "" Then Exit Sub
On Error GoTo ERROR::
Open App.Path & "\Levels\" & FileName & ".til" For Random As #FileNum
    For c = 1 To 20
    For d = 1 To 10
        Put #FileNum, ((c - 1) * 20) + d, MemBoard.Tiles(c, d)
    Next
    Next
Close #FileNumber
Exit Sub
ERROR::
MsgBox ("Error: Unable To Save")
End Sub

Private Sub Menu_Edit_Tiles_Blank_Click()
ResetTileMenu
Menu_Edit_Tiles_Blank.Checked = True
End Sub

Private Sub Menu_Edit_Tiles_Cross_Click()
ResetTileMenu
Menu_Edit_Tiles_Cross.Checked = True
End Sub

Private Sub Menu_Edit_Tiles_Naught_Click()
ResetTileMenu
Menu_Edit_Tiles_Naught.Checked = True
End Sub

Private Sub Menu_Edit_Tiles_None_Click()
ResetTileMenu
Menu_Edit_Tiles_None.Checked = True
End Sub

Private Sub Menu_Game_Edit_Click()
Menu_Edit.Enabled = True
Menu_Game.Enabled = False
MemClearBoard
Cls
WriteBoard
UserMode = 1
End Sub

Private Sub Menu_Game_Load_Click()
PlayScreen.Enabled = False
LevelLoader.Show
End Sub

Private Sub Menu_Game_Names_Click()
Player1 = InputBox("Naughts: Please Enter Your Name", "Naughtz - Player Setup", Player1)
Player2 = InputBox("Crosses: Please Enter Your Name", "Naughtz - Player Setup", Player2)
If Player1 = "" Then Player1 = "Naughts"
If Player2 = "" Then Player2 = "Crosses"
End Sub

Private Sub Menu_Game_New_Click()
NewGame
End Sub

Private Sub Menu_Game_Options_Click()
OptionMenu.Show
PlayScreen.Enabled = False
End Sub

Private Sub Menu_Game_Preset_Full_Click()
ClearMemBoard
MemClearBoard
Cls
WriteBoard
End Sub

Private Sub Menu_Game_Preset_Normal_Click()
ClearMemStandardBoard
MemClearBoard
Cls
WriteBoard
End Sub

Private Sub Menu_Game_Switch_Click()
If MsgBox("Switch Players?", vbYesNo) = vbNo Then Exit Sub
If PlayerLogo(0).Visible = True Then
    PlayerLogo(0).Visible = False
    PlayerLogo(1).Visible = True
Else
    PlayerLogo(1).Visible = False
    PlayerLogo(0).Visible = True
End If
End Sub

Private Sub Menu_Quit_Click()
End
End Sub

Private Sub UFOMove_Timer()
If UFOType = UFOGreen Then
    GreenUFO.Left = GreenUFO.Left + 10
    If GreenUFO.Left = Board.Tiles(UFOAttackX, UFOAttackY).XPos - 10 Then
        If Rnd > 0.5 Then
            Board.Tiles(UFOAttackX, UFOAttackY).Tile = Cross
        Else
            Board.Tiles(UFOAttackX, UFOAttackY).Tile = Naught
        End If
        WriteBoard
    End If
    If GreenUFO.Left > BoardLeft + 440 Then
        UFOMove.Enabled = False
        UFOInUse = False
        GreenUFO.Visible = False
    End If
End If
If UFOType = UFOBlue Then
    BlueUFO.Left = BlueUFO.Left + 10
    If BlueUFO.Left = Board.Tiles(UFOAttackX, UFOAttackY).XPos - 10 Then
        Board.Tiles(UFOAttackX, UFOAttackY).Tile = Normal
        WriteBoard
    End If
    If BlueUFO.Left > BoardLeft + 440 Then
        UFOMove.Enabled = False
        UFOInUse = False
        BlueUFO.Visible = False
    End If
End If
If UFOType = UFORed Then
    RedUFO.Left = RedUFO.Left + 10
    If RedUFO.Left > (BoardLeft - 20) And RedUFO.Left < (BoardLeft + 380) And (((RedUFO.Left + 10) - BoardLeft) Mod 20) = 0 Then
        If Board.Tiles(((RedUFO.Left + 10) - BoardLeft) / 20, UFOAttackY).Tile <> Missing Then
            Board.Tiles(((RedUFO.Left + 10) - BoardLeft) / 20, UFOAttackY).Tile = Normal
            WriteBoard
        End If
    End If
    If RedUFO.Left > BoardLeft + 440 Then
        UFOMove.Enabled = False
        UFOInUse = False
        RedUFO.Visible = False
    End If
End If
End Sub
