Attribute VB_Name = "Declarations"
' Constants

Public Const BoardTop As Integer = 125
Public Const BoardLeft As Integer = 100

' Tiles
Public Const Normal = 0
Public Const Missing = 1
Public Const Naught = 2
Public Const Cross = 3
Public Const Teleport = 4
Public Const Slow = 5

Public Const SlowDelay = 2

' Player
Public Const None = 0
Public Const PNaught = 1
Public Const PCross = 2

' UFO Types
Public Const UFOGreen = 1
Public Const UFOBlue = 2
Public Const UFORed = 3

' UFO Limits
Public Const GreenLimit = 0
Public Const BlueLimit = 0.4
Public Const RedLimit = 0.8
Public Const UFOOccurance = 0.5


' User-defined Types

Public Type TileType
    XPos As Integer
    YPos As Integer
    Tile As Integer
    Player As Integer
    Delay As Integer
End Type

Public Type BoardType
    Title As String
    Tiles(20, 10) As TileType
End Type


' Variables

Public Board As BoardType
Public MemBoard As BoardType
Public UserMode As Integer
Public Player1 As String
Public Player2 As String
Public RegisterAbout As Boolean
Public SignID As String

Public AllowUFO As Boolean

Public UFOAttackX As Integer
Public UFOAttackY As Integer
Public UFOInUse As Boolean
Public UFOType As Integer
Public CanNotPlay As Boolean


' Functions

Function ClearBoard()
    For a = 1 To 20
        For b = 1 To 10
            Board.Tiles(a, b).Player = None
            Board.Tiles(a, b).Delay = 0
            Board.Tiles(a, b).XPos = BoardLeft + ((a - 1) * 20)
            Board.Tiles(a, b).YPos = BoardTop + ((b - 1) * 20)
            Board.Tiles(a, b).Tile = Normal
        Next
    Next
End Function

Function MemClearBoard()
    Board = MemBoard
End Function

Function ClearMemBoard()
    For a = 1 To 20
        For b = 1 To 10
            MemBoard.Tiles(a, b).Player = None
            MemBoard.Tiles(a, b).Delay = 0
            MemBoard.Tiles(a, b).XPos = BoardLeft + ((a - 1) * 20)
            MemBoard.Tiles(a, b).YPos = BoardTop + ((b - 1) * 20)
            MemBoard.Tiles(a, b).Tile = Normal
        Next
    Next
End Function

Function ClearMemBlankBoard()
    For a = 1 To 20
        For b = 1 To 10
            MemBoard.Tiles(a, b).Player = None
            MemBoard.Tiles(a, b).Delay = 0
            MemBoard.Tiles(a, b).XPos = BoardLeft + ((a - 1) * 20)
            MemBoard.Tiles(a, b).YPos = BoardTop + ((b - 1) * 20)
            MemBoard.Tiles(a, b).Tile = Missing
        Next
    Next
End Function

Function ClearMemStandardBoard()
    For a = 1 To 20
        For b = 1 To 10
            MemBoard.Tiles(a, b).Player = None
            MemBoard.Tiles(a, b).Delay = 0
            MemBoard.Tiles(a, b).XPos = BoardLeft + ((a - 1) * 20)
            MemBoard.Tiles(a, b).YPos = BoardTop + ((b - 1) * 20)
            MemBoard.Tiles(a, b).Tile = Missing
        Next
    Next
    MemBoard.Tiles(9, 4).Tile = Normal
    MemBoard.Tiles(10, 4).Tile = Normal
    MemBoard.Tiles(11, 4).Tile = Normal
    MemBoard.Tiles(9, 5).Tile = Normal
    MemBoard.Tiles(10, 5).Tile = Normal
    MemBoard.Tiles(11, 5).Tile = Normal
    MemBoard.Tiles(9, 6).Tile = Normal
    MemBoard.Tiles(10, 6).Tile = Normal
    MemBoard.Tiles(11, 6).Tile = Normal
End Function

Function DrawTile(X As Integer, Y As Integer, DrawBoard As BoardType)
    If DrawBoard.Tiles(X, Y).Tile = Normal Then
        PlayScreen.PaintPicture LoadResPicture("BLANKTILE", 0), DrawBoard.Tiles(X, Y).XPos, DrawBoard.Tiles(X, Y).YPos
    ElseIf DrawBoard.Tiles(X, Y).Tile = Naught Then
        PlayScreen.PaintPicture LoadResPicture("NAUGHT", 0), DrawBoard.Tiles(X, Y).XPos, DrawBoard.Tiles(X, Y).YPos
    ElseIf DrawBoard.Tiles(X, Y).Tile = Cross Then
        PlayScreen.PaintPicture LoadResPicture("CROSS", 0), DrawBoard.Tiles(X, Y).XPos, DrawBoard.Tiles(X, Y).YPos
    End If
End Function

Function WriteBoard()
Dim a As Integer, b As Integer
For a = 1 To 20
    For b = 1 To 10
        DrawTile a, b, Board
    Next
Next
End Function

Function NaughtWin(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    CanNotPlay = True
    Dim XPosOne As Single
    Dim YPosOne As Single
    Dim XPosTwo As Single
    Dim YPosTwo As Single
    
    XPosOne = CSng(X1) - 1
    YPosOne = CSng(Y1) - 1
    XPosTwo = CSng(X2) - 1
    YPosTwo = CSng(Y2) - 1
    XPosOne = XPosOne * 20
    YPosOne = YPosOne * 20
    XPosTwo = XPosTwo * 20
    YPosTwo = YPosTwo * 20
    
    XPosOne = XPosOne + BoardLeft
    YPosOne = YPosOne + BoardTop
    XPosTwo = XPosTwo + BoardLeft
    YPosTwo = YPosTwo + BoardTop
    
    Beep
    PlayScreen.Line (XPosOne + 10, YPosOne + 10)-(XPosTwo + 10, YPosTwo + 10), RGB(0, 0, 255)
    
    With PlayScreen
        .PlayerLogo(0).Visible = False
        .PlayerLogo(1).Visible = False
        Pause 0.5
        .PlayerLogo(0).Visible = True
        Pause 0.5
        .PlayerLogo(0).Visible = False
        Pause 0.5
        .PlayerLogo(0).Visible = True
        Pause 0.5
        .PlayerLogo(0).Visible = False
        Pause 0.5
        .PlayerLogo(0).Visible = True
        If Player1 = "Naughts" Then
            MsgBox ("Naughts Win!")
        Else
            MsgBox (Player1 & " Wins!")
        End If
    End With
    MemClearBoard
    PlayScreen.Cls
    WriteBoard
    CanNotPlay = False
End Function

Function CrossWin(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    CanNotPlay = True
    Dim XPosOne As Single
    Dim YPosOne As Single
    Dim XPosTwo As Single
    Dim YPosTwo As Single
    
    XPosOne = CSng(X1) - 1
    YPosOne = CSng(Y1) - 1
    XPosTwo = CSng(X2) - 1
    YPosTwo = CSng(Y2) - 1
    XPosOne = XPosOne * 20
    YPosOne = YPosOne * 20
    XPosTwo = XPosTwo * 20
    YPosTwo = YPosTwo * 20
    
    XPosOne = XPosOne + BoardLeft
    YPosOne = YPosOne + BoardTop
    XPosTwo = XPosTwo + BoardLeft
    YPosTwo = YPosTwo + BoardTop
    
    Beep
    PlayScreen.Line (XPosOne + 10, YPosOne + 10)-(XPosTwo + 10, YPosTwo + 10), RGB(255, 0, 0)
    
    With PlayScreen
        .PlayerLogo(0).Visible = False
        .PlayerLogo(1).Visible = False
        Pause 0.5
        .PlayerLogo(1).Visible = True
        Pause 0.5
        .PlayerLogo(1).Visible = False
        Pause 0.5
        .PlayerLogo(1).Visible = True
        Pause 0.5
        .PlayerLogo(1).Visible = False
        Pause 0.5
        .PlayerLogo(1).Visible = True
        If Player2 = "Crosses" Then
            MsgBox ("Crosses Win!")
        Else
            MsgBox (Player2 & " Wins!")
        End If
    End With
    MemClearBoard
    PlayScreen.Cls
    WriteBoard
    CanNotPlay = False
End Function

Function CheckPlayerGo(TileX As Integer, TileY As Integer, PlayerChange As Boolean) As Boolean
Dim Check(5, 5) As Boolean
CheckPlayerGo = False
    If PlayScreen.PlayerLogo(0).Visible = True Then
    Board.Tiles(TileX, TileY).Tile = Naught
    If PlayerChange = True Then PlayScreen.PlayerLogo(0).Visible = False
    If PlayerChange = True Then PlayScreen.PlayerLogo(1).Visible = True
    Else
    Board.Tiles(TileX, TileY).Tile = Cross
    If PlayerChange = True Then PlayScreen.PlayerLogo(1).Visible = False
    If PlayerChange = True Then PlayScreen.PlayerLogo(0).Visible = True
    End If
    
    DrawTile TileX, TileY, Board
    
    For a = 1 To 5
        For b = 1 To 5
            Check(a, b) = False
        Next
    Next
    
    On Error Resume Next
    
    If Board.Tiles(TileX, TileY).Tile = Naught Then
    
    If Not ((TileX - 2) < 1 Or (TileY - 2) < 1) Then If Board.Tiles(TileX - 2, TileY - 2).Tile = Naught Then Check(1, 1) = True
    If Not ((TileX - 1) < 1 Or (TileY - 2) < 1) Then If Board.Tiles(TileX - 1, TileY - 2).Tile = Naught Then Check(2, 1) = True
    If Not ((TileX) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX, TileY - 2).Tile = Naught Then Check(3, 1) = True
    If Not ((TileX + 1) > 20 Or (TileY - 2) < 1) Then If Board.Tiles(TileX + 1, TileY - 2).Tile = Naught Then Check(4, 1) = True
    If Not ((TileX + 2) > 20 Or (TileY - 2) < 1) Then If Board.Tiles(TileX + 2, TileY - 2).Tile = Naught Then Check(5, 1) = True
    If Not ((TileX - 2) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX - 2, TileY - 1).Tile = Naught Then Check(1, 2) = True
    If Not ((TileX - 1) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX - 1, TileY - 1).Tile = Naught Then Check(2, 2) = True
    If Not ((TileX) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX, TileY - 1).Tile = Naught Then Check(3, 2) = True
    If Not ((TileX + 1) > 20 Or (TileY - 1) < 1) Then If Board.Tiles(TileX + 1, TileY - 1).Tile = Naught Then Check(4, 2) = True
    If Not ((TileX + 2) > 20 Or (TileY - 1) < 1) Then If Board.Tiles(TileX + 2, TileY - 1).Tile = Naught Then Check(5, 2) = True
    If Not ((TileX - 2) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX - 2, TileY).Tile = Naught Then Check(1, 3) = True
    If Not ((TileX - 1) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX - 1, TileY).Tile = Naught Then Check(2, 3) = True
    If Not ((TileX) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX, TileY).Tile = Naught Then Check(3, 3) = True
    If Not ((TileX + 1) > 20 Or (TileY) < 1) Then If Board.Tiles(TileX + 1, TileY).Tile = Naught Then Check(4, 3) = True
    If Not ((TileX + 2) > 20 Or (TileY) < 1) Then If Board.Tiles(TileX + 2, TileY).Tile = Naught Then Check(5, 3) = True
    If Not ((TileX - 2) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX - 2, TileY + 1).Tile = Naught Then Check(1, 4) = True
    If Not ((TileX - 1) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX - 1, TileY + 1).Tile = Naught Then Check(2, 4) = True
    If Not ((TileX) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX, TileY + 1).Tile = Naught Then Check(3, 4) = True
    If Not ((TileX + 1) > 20 Or (TileY + 1) > 10) Then If Board.Tiles(TileX + 1, TileY + 1).Tile = Naught Then Check(4, 4) = True
    If Not ((TileX + 2) > 20 Or (TileY + 1) > 10) Then If Board.Tiles(TileX + 2, TileY + 1).Tile = Naught Then Check(5, 4) = True
    If Not ((TileX - 2) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX - 2, TileY + 2).Tile = Naught Then Check(1, 5) = True
    If Not ((TileX - 1) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX - 1, TileY + 2).Tile = Naught Then Check(2, 5) = True
    If Not ((TileX) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX, TileY + 2).Tile = Naught Then Check(3, 5) = True
    If Not ((TileX + 1) > 20 Or (TileY + 2) > 10) Then If Board.Tiles(TileX + 1, TileY + 2).Tile = Naught Then Check(4, 5) = True
    If Not ((TileX + 2) > 20 Or (TileY + 2) > 10) Then If Board.Tiles(TileX + 2, TileY + 2).Tile = Naught Then Check(5, 5) = True
    
    End If
    
    If Board.Tiles(TileX, TileY).Tile = Cross Then
    
    If Not ((TileX - 2) < 1 Or (TileY - 2) < 1) Then If Board.Tiles(TileX - 2, TileY - 2).Tile = Cross Then Check(1, 1) = True
    If Not ((TileX - 1) < 1 Or (TileY - 2) < 1) Then If Board.Tiles(TileX - 1, TileY - 2).Tile = Cross Then Check(2, 1) = True
    If Not ((TileX) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX, TileY - 2).Tile = Cross Then Check(3, 1) = True
    If Not ((TileX + 1) > 20 Or (TileY - 2) < 1) Then If Board.Tiles(TileX + 1, TileY - 2).Tile = Cross Then Check(4, 1) = True
    If Not ((TileX + 2) > 20 Or (TileY - 2) < 1) Then If Board.Tiles(TileX + 2, TileY - 2).Tile = Cross Then Check(5, 1) = True
    If Not ((TileX - 2) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX - 2, TileY - 1).Tile = Cross Then Check(1, 2) = True
    If Not ((TileX - 1) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX - 1, TileY - 1).Tile = Cross Then Check(2, 2) = True
    If Not ((TileX) < 1 Or (TileY - 1) < 1) Then If Board.Tiles(TileX, TileY - 1).Tile = Cross Then Check(3, 2) = True
    If Not ((TileX + 1) > 20 Or (TileY - 1) < 1) Then If Board.Tiles(TileX + 1, TileY - 1).Tile = Cross Then Check(4, 2) = True
    If Not ((TileX + 2) > 20 Or (TileY - 1) < 1) Then If Board.Tiles(TileX + 2, TileY - 1).Tile = Cross Then Check(5, 2) = True
    If Not ((TileX - 2) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX - 2, TileY).Tile = Cross Then Check(1, 3) = True
    If Not ((TileX - 1) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX - 1, TileY).Tile = Cross Then Check(2, 3) = True
    If Not ((TileX) < 1 Or (TileY) < 1) Then If Board.Tiles(TileX, TileY).Tile = Cross Then Check(3, 3) = True
    If Not ((TileX + 1) > 20 Or (TileY) < 1) Then If Board.Tiles(TileX + 1, TileY).Tile = Cross Then Check(4, 3) = True
    If Not ((TileX + 2) > 20 Or (TileY) < 1) Then If Board.Tiles(TileX + 2, TileY).Tile = Cross Then Check(5, 3) = True
    If Not ((TileX - 2) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX - 2, TileY + 1).Tile = Cross Then Check(1, 4) = True
    If Not ((TileX - 1) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX - 1, TileY + 1).Tile = Cross Then Check(2, 4) = True
    If Not ((TileX) < 1 Or (TileY + 1) > 10) Then If Board.Tiles(TileX, TileY + 1).Tile = Cross Then Check(3, 4) = True
    If Not ((TileX + 1) > 20 Or (TileY + 1) > 10) Then If Board.Tiles(TileX + 1, TileY + 1).Tile = Cross Then Check(4, 4) = True
    If Not ((TileX + 2) > 20 Or (TileY + 1) > 10) Then If Board.Tiles(TileX + 2, TileY + 1).Tile = Cross Then Check(5, 4) = True
    If Not ((TileX - 2) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX - 2, TileY + 2).Tile = Cross Then Check(1, 5) = True
    If Not ((TileX - 1) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX - 1, TileY + 2).Tile = Cross Then Check(2, 5) = True
    If Not ((TileX) < 1 Or (TileY + 2) > 10) Then If Board.Tiles(TileX, TileY + 2).Tile = Cross Then Check(3, 5) = True
    If Not ((TileX + 1) > 20 Or (TileY + 2) > 10) Then If Board.Tiles(TileX + 1, TileY + 2).Tile = Cross Then Check(4, 5) = True
    If Not ((TileX + 2) > 20 Or (TileY + 2) > 10) Then If Board.Tiles(TileX + 2, TileY + 2).Tile = Cross Then Check(5, 5) = True
    
    
    End If
    
    On Error GoTo 0
    
    If Board.Tiles(TileX, TileY).Tile = Naught Then
    
    If Check(1, 1) = True And Check(2, 2) = True And Check(3, 3) = True Then
        NaughtWin TileX - 2, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(2, 2) = True And Check(3, 3) = True And Check(4, 4) = True Then
        NaughtWin TileX - 1, TileY - 1, TileX + 1, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(4, 4) = True And Check(5, 5) = True Then
        NaughtWin TileX, TileY, TileX + 2, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(5, 1) = True And Check(4, 2) = True And Check(3, 3) = True Then
        NaughtWin TileX + 2, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(4, 2) = True And Check(3, 3) = True And Check(2, 4) = True Then
        NaughtWin TileX + 1, TileY - 1, TileX - 1, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(2, 4) = True And Check(1, 5) = True Then
        NaughtWin TileX, TileY, TileX - 2, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(1, 3) = True And Check(2, 3) = True And Check(3, 3) = True Then
        NaughtWin TileX - 2, TileY, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(2, 3) = True And Check(3, 3) = True And Check(4, 3) = True Then
        NaughtWin TileX - 1, TileY, TileX + 1, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(4, 3) = True And Check(5, 3) = True Then
        NaughtWin TileX, TileY, TileX + 2, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(3, 1) = True And Check(3, 2) = True And Check(3, 3) = True Then
        NaughtWin TileX, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 2) = True And Check(3, 3) = True And Check(3, 4) = True Then
        NaughtWin TileX, TileY - 1, TileX, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(3, 4) = True And Check(3, 5) = True Then
        NaughtWin TileX, TileY, TileX, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    End If
    
    
    If Board.Tiles(TileX, TileY).Tile = Cross Then
    
    If Check(1, 1) = True And Check(2, 2) = True And Check(3, 3) = True Then
        CrossWin TileX - 2, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(2, 2) = True And Check(3, 3) = True And Check(4, 4) = True Then
        CrossWin TileX - 1, TileY - 1, TileX + 1, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(4, 4) = True And Check(5, 5) = True Then
        CrossWin TileX, TileY, TileX + 2, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(5, 1) = True And Check(4, 2) = True And Check(3, 3) = True Then
        CrossWin TileX + 2, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(4, 2) = True And Check(3, 3) = True And Check(2, 4) = True Then
        CrossWin TileX + 1, TileY - 1, TileX - 1, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(2, 4) = True And Check(1, 5) = True Then
        CrossWin TileX, TileY, TileX - 2, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(1, 3) = True And Check(2, 3) = True And Check(3, 3) = True Then
        CrossWin TileX - 2, TileY, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(2, 3) = True And Check(3, 3) = True And Check(4, 3) = True Then
        CrossWin TileX - 1, TileY, TileX + 1, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(4, 3) = True And Check(5, 3) = True Then
        CrossWin TileX, TileY, TileX + 2, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    If Check(3, 1) = True And Check(3, 2) = True And Check(3, 3) = True Then
        CrossWin TileX, TileY - 2, TileX, TileY
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 2) = True And Check(3, 3) = True And Check(3, 4) = True Then
        CrossWin TileX, TileY - 1, TileX, TileY + 1
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    If Check(3, 3) = True And Check(3, 4) = True And Check(3, 5) = True Then
        CrossWin TileX, TileY, TileX, TileY + 2
        CheckPlayerGo = True
        GoTo EndCheck::
    End If
    
    End If
    
    ' Check for full board (new in 1.1)
    If CheckFull = True Then
        MsgBox ("Draw")
        NewGame
    End If
EndCheck::
End Function

Function ResetTileMenu()
    PlayScreen.Menu_Edit_Tiles_Blank.Checked = False
    PlayScreen.Menu_Edit_Tiles_None.Checked = False
    PlayScreen.Menu_Edit_Tiles_Naught.Checked = False
    PlayScreen.Menu_Edit_Tiles_Cross.Checked = False
End Function

Function Pause(Delay As Single)
    Dim CS As Single
    CS = Timer
    Do Until CS + Delay < Timer
    DoEvents
    Loop
End Function

Function CheckFull() As Boolean
    Dim X As Long, Y As Long
    CheckFull = True
    For X = 1 To 20
        For Y = 1 To 10
            If Board.Tiles(X, Y).Tile = Normal Then CheckFull = False
        Next
    Next
End Function

Function NewGame()
' Function moved from PlayScreen.frm in 1.1
PlayScreen.Cls
MemClearBoard
WriteBoard
End Function
