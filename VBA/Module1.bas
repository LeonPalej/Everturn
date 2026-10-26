Attribute VB_Name = "Module1"
' --------------------------------------------------------------------------------------------- API Hook for reading key presses from windows
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' --------------------------------------------------------------------------------------------- Keyboard Keys (WASD)
Const VK_W As Long = &H57
Const VK_S As Long = &H53
Const VK_A As Long = &H41
Const VK_D As Long = &H44
Const VK_ESCAPE As Long = &H1B

' --------------------------------------------------------------------------------------------- Variables

Public Ws As Worksheet

Public Decore As Shape
Public StartButton As Shape
Public Count As Shape
Public Player As Shape
Public ScoreBoard As Shape
Public BestScore As Shape
Public CoinShape As Shape
Public Controls As Shape
Public Mes As Shape
Public GameRunning As Boolean
Public DeletePixelX As Double
Public SpawnPixelX As Double
Public DangerPixelX As Double
Public Coins As Collection
Private ColCounter As Long

Public Const PlayerRow As Long = 21
Public Const PlayerCol As Long = 175
Public Const RowLenght As Integer = 24
Public Const ColLenght As Integer = 24
Public Const DuckCellTick As Double = 1       ' How much cells the duck moves every tick
Public Const MazeSlots As Long = 15           ' Height in blocks
Public Const MazeCellTick As Double = 1       ' How much cells the maze moves every tick
Public Const WallDensity As Double = 0.48     ' Wall / Air ratio
Public Const MazeTopRow As Long = 3           ' Row where the maze starts
Public Const BlockCells As Long = 3           ' How many cells is one block n x n
Public Const MaxTurn As Long = 4              ' max blocks the safe path can go up/down per column
Public Const DeleteCol As Long = 110          ' Maze end
Public Const SpawnCol As Long = 240           ' Maze start
Public Const DangerCol As Long = 119          ' Danger col
Public Const HitboxScaleX As Double = 0.6     ' % of the players width for the hitbox
Public Const HitboxScaleY As Double = 0.6     ' % of the players height for the hitbox
Public Const ATotal As Double = -2            ' total duck movement on SCREEN per tick while holding A (cells, includes the drift)
Public Const DTotal As Double = 1             ' total duck movement on SCREEN per tick while holding D (cells, includes the drift)
Public Const TickMs As Long = 25              ' smothnes
Public Const BaseTickMs As Long = 50          ' How long is one tick, 1000 / value = TPS, Higher TPS gives a smoother game at the same speed
Public Const ClearBehind As Long = 3          ' Clear Col behind the player at the start
Public Const ClearAhead As Long = 3           ' Clear Col in front of the player at the start
Public Const CoinValue As Long = 50           ' Value of one coin
Public Const CoinEveryCols As Long = 5
Public Const CoinOnSafePath As Boolean = False
Public Const CoinSize As Double = 40
Public Const StartGameSpeed As Double = 0.5
Public Const SpeedUpEveryCols As Long = 20
Public Const SpeedMultiplier As Double = 1.03
Public Const MaxGameSpeed As Double = 1.5
Public Const SkinCount As Long = 5
Public Const CoinSkinCount As Long = 2
Public Const PlayerSize As Double = 0
Public BestColor As Long
Public ScoreColor As Long
Public MazeColor As Long
Public SecColorB As Long
Public SecColorA As Long
Public PriColor As Long
Public Const RightWallCol As Long = 237     ' IC: the duck can't go past the left edge of this column
Public RightWallPixelX As Double
' :) IT WORKING, FINALY, I CAN SLEAP NOW

Private Const MaxStepPx As Double = 8
Private Const MaxPushPx As Long = 100
Public Const BlockW = ColLenght * BlockCells
Public Const BlockH = RowLenght * BlockCells

Private PathSlot As Long
' --------------------------------------------------------------------------------------------- Settings
Sub Coll()
    Set Ws = ActiveSheet
    Call SetVariables
    Ws.Range("DO3:IB48").Interior.Color = PriColor
    Ws.Range("DO2:IB2").Interior.Color = SecColorB
    Controls.Fill.ForeColor.RGB = SecColorA
    Controls.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PriColor
    Decore.Fill.ForeColor.RGB = SecColorA
    Mes.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = SecColorB
    StartButton.Fill.ForeColor.RGB = SecColorB
    StartButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    Mes.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    ScoreBoard.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    BestScore.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BestColor
    Count.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
End Sub

' --------------------------------------------------------------------------------------------- Assets
Function GetSelectedCoin() As Long
    Dim n As Long
    n = Val(ActiveSheet.Shapes("coinframe").AlternativeText)
    If n < 1 Or n > CoinSkinCount Then n = 1
    GetSelectedCoin = n
    
End Function

Sub SelectCoin(ByVal n As Long)
    Dim fr As Shape, sk As Shape
    If GameRunning Then Exit Sub
    
    Set fr = ActiveSheet.Shapes("coinframe")
    Set sk = ActiveSheet.Shapes("coinskin" & n)
    
    fr.AlternativeText = CStr(n)
    fr.Left = sk.Left - 4
    fr.Top = sk.Top - 4
    fr.Width = sk.Width + 8
    fr.Height = sk.Height + 8
    fr.Visible = msoTrue
End Sub

Function TrailingNumber(ByVal nm As String) As Long
    Dim i As Long, digits As String
    For i = Len(nm) To 1 Step -1
        If Mid(nm, i, 1) Like "[0-9]" Then digits = Mid(nm, i, 1) & digits Else Exit For
    Next i
    TrailingNumber = Val(digits)
End Function

Sub PickCoin()
    Dim n As Long
    On Error Resume Next
    n = TrailingNumber(Application.Caller)
    On Error GoTo 0
    If n >= 1 And n <= CoinSkinCount Then SelectCoin n
End Sub

Function GetSelectedSkin() As Long
    Dim n As Long
    n = Val(ActiveSheet.Shapes("skinframe").AlternativeText)
    If n < 1 Or n > SkinCount Then n = 1
    GetSelectedSkin = n
    
End Function

Sub MakePlayerCopy()
    Dim n As Long
    Dim sh As Shape
    
    ActiveSheet.Unprotect
    n = GetSelectedSkin()
    On Error Resume Next
    ActiveSheet.Shapes("player").Delete
    On Error GoTo 0
    
    ActiveSheet.Shapes("skin" & n).Duplicate
    Set sh = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)
    sh.Name = "player"
    sh.OnAction = ""
    If PlayerSize > 0 Then
        sh.LockAspectRatio = msoTrue
        sh.Width = PlayerSize
    End If
    sh.Left = PlayerCol * ColLenght
    sh.Top = PlayerRow * RowLenght
    sh.Visible = msoTrue
    sh.ZOrder msoBringToFront
End Sub
    
    
    ' Remember the choice, move the frame, and show the new duck at the start position
Sub SelectSkin(ByVal n As Long)
    Dim fr As Shape, sk As Shape
    If GameRunning Then Exit Sub

    Set fr = ActiveSheet.Shapes("skinframe")
    Set sk = ActiveSheet.Shapes("skin" & n)

    fr.AlternativeText = CStr(n)
    fr.Left = sk.Left - 4
    fr.Top = sk.Top - 4
    fr.Width = sk.Width + 8
    fr.Height = sk.Height + 8
    fr.Visible = msoTrue

    MakePlayerCopy
End Sub

' ONE macro for all five buttons: it reads the number from the button's name (btn3 -> 3)
Sub PickSkin()
    Dim n As Long
    On Error Resume Next
    n = TrailingNumber(Application.Caller)
    On Error GoTo 0
    If n >= 1 And n <= SkinCount Then SelectSkin n
End Sub
' --------------------------------------------------------------------------------------------- Program Scripts
Sub SetVariables()
    Set Decore = Application.ActiveSheet.Shapes("decore")
    Set StartButton = Application.ActiveSheet.Shapes("startbutton")
    If Len(ActiveSheet.Shapes("skinframe").AlternativeText) > 0 Then
        Call MakePlayerCopy                    ' a skin was chosen: use a fresh copy of it
    Else
        With ActiveSheet.Shapes("player")      ' nothing chosen yet: keep the duck on screen, just teleport it
            .Left = PlayerCol * ColLenght
            .Top = PlayerRow * RowLenght
        End With
    End If
    Set Player = Application.ActiveSheet.Shapes("player")
    Set CoinShape = Application.ActiveSheet.Shapes("coinskin" & GetSelectedCoin())
    Set ScoreBoard = Application.ActiveSheet.Shapes("scoreboard")
    Set BestScore = Application.ActiveSheet.Shapes("bestscore")
    Set Count = Application.ActiveSheet.Shapes("count")
    Set Controls = Application.ActiveSheet.Shapes("controls")
    Set Mes = Application.ActiveSheet.Shapes("mes")
    DeletePixelX = Ws.Cells(1, DeleteCol).Left
    SpawnPixelX = Ws.Cells(1, SpawnCol).Left
    DangerPixelX = Ws.Cells(1, DangerCol).Left
    RightWallPixelX = Ws.Cells(1, RightWallCol).Left
    GameRunning = False
    BestColor = Ws.Range("CM8").Interior.Color
    ScoreColor = Ws.Range("CM9").Interior.Color
    MazeColor = Ws.Range("CM10").Interior.Color
    SecColorB = Ws.Range("CM11").Interior.Color
    SecColorA = Ws.Range("CM12").Interior.Color
    PriColor = Ws.Range("CM13").Interior.Color
    
End Sub

' Wait that dose not stop events; time in ms
Sub Wait(MilliSeconds As Long)
    Dim EndTime  As Double
    EndTime = Timer + (MilliSeconds / 1000)
    Do While Timer < EndTime
        DoEvents
        
    Loop
End Sub
' --------------------------------------------------------------------------------------------- Animation Scripts


' Moving a shape; Shape to move, how many rows, how many columns, Speed of the move, Easing curve
Sub MoveShape(a As Shape, Trow As Integer, Tcol As Integer, Speed As Long, ModifLevel As Integer)
    Dim Stpc As Integer
    Dim TotalSteps As Long
    Dim RowStep As Double
    Dim ColStep As Double
    Dim Progress As Double
    Dim StartLeft As Double
    Dim StartTop As Double
    Dim Modifier As Double

    Stpc = 20    ' Steps Per Cell

    TotalSteps = Application.WorksheetFunction.Max(Abs(Trow), Abs(Tcol)) * Stpc
    ColStep = (ColLenght * Tcol)
    RowStep = (RowLenght * Trow)
    StartLeft = a.Left
    StartTop = a.Top

    For i = 1 To TotalSteps

        Progress = i / TotalSteps
        Modifier = (1 - Cos(Progress * 3.14159265358979)) / 2
        For m = 1 To ModifLevel - 1
            Modifier = Modifier * Modifier * (3 - 2 * Modifier)
        Next m
        a.Left = StartLeft + (ColStep * Modifier)
        a.Top = StartTop + (RowStep * Modifier)

        Wait Speed

    Next i

End Sub


' Fadeing out a shape; Shape to Fade, How fast, Fade fill color, Fade Line color
Sub FadeOutShape(a As Shape, Speed As Long, FillY As Boolean, LineY As Boolean)
    Dim TotalSteps As Long
    Dim Progress As Double

    TotalSteps = 20

    For i = 1 To TotalSteps

        Progress = i / TotalSteps
        If FillY Then
            a.Fill.Transparency = Progress
        End If
        If LineY Then
            a.Line.Transparency = Progress
        End If
        Wait Speed
    Next i
    a.Visible = msoFalse
End Sub


' Fadeing in a shape; Shape to Fade, How fast, Fade fill color, Fade Line color
Sub FadeInShape(a As Shape, Speed As Long, FillY As Boolean, LineY As Boolean)
    Dim TotalSteps As Long
    Dim Progress As Double
    Dim i As Long

    a.Visible = msoTrue
    TotalSteps = 20

    For i = 1 To TotalSteps

        Progress = i / TotalSteps
        If FillY Then
            a.Fill.Transparency = 1 - Progress
        End If
        If LineY Then
            a.Line.Transparency = 1 - Progress
        End If
        Wait Speed
    Next i

End Sub


' --------------------------------------------------------------------------------------------- Game Scripts
Sub StartEnd(UpdateStateTo As String)
    If UpdateStateTo = "Start" Then
        Application.OnKey "w", ""
        Application.OnKey "s", ""
        Application.OnKey "a", ""
        Application.OnKey "d", ""
        Application.OnKey "W", ""
        Application.OnKey "S", ""
        Application.OnKey "A", ""
        Application.OnKey "D", ""
        ActiveWindow.Zoom = 50
        ActiveWindow.ScrollColumn = Ws.Range("DO22").Column
        ActiveWindow.ScrollRow = 2
        Ws.ScrollArea = "DO22:IB48"
        Ws.Protect UserInterfaceOnly:=True
        Ws.EnableSelection = xlNoSelection
        Call FadeOutShape(StartButton, 40, True, True)
        Call FadeOutShape(Controls, 40, True, True)
        Call FadeOutShape(Mes, 40, False, True)
        Call MoveShape(Decore, 35, 0, 4, 2)
        Count.TextFrame.Characters.Text = "3"
        Wait 1000
        Count.TextFrame.Characters.Text = "2"
        Wait 1000
        Count.TextFrame.Characters.Text = "1"
        Wait 1000
        Count.TextFrame.Characters.Text = ""
        
    ElseIf UpdateStateTo = "End" Then
        Application.OnKey "w"
        Application.OnKey "s"
        Application.OnKey "a"
        Application.OnKey "d"
        Application.OnKey "W"
        Application.OnKey "S"
        Application.OnKey "A"
        Application.OnKey "D"
        Call FadeOutShape(Player, 40, False, True)
        Player.Left = PlayerCol * ColLenght
        Player.Top = PlayerRow * RowLenght
        Call MoveShape(Decore, -35, 0, 4, 2)
        Call FadeInShape(StartButton, 40, True, True)
        Call FadeInShape(Player, 40, False, True)
        Call FadeInShape(Mes, 40, False, True)
        Call FadeInShape(Controls, 40, True, True)
        Ws.Unprotect
        Ws.ScrollArea = ""
        Ws.EnableSelection = xlNoRestrictions
        
    End If
End Sub



Sub SpawnMazeColumn(ByVal x As Double, ByVal ForceOpen As Boolean, ByVal Blocks As Collection)
    Dim isWall() As Boolean
    ReDim isWall(1 To MazeSlots)
    Dim s As Long
    Dim NewSlot As Long
    Dim lo As Long
    Dim hi As Long
    Dim RunStart As Long
    Dim IsW As Boolean
    Dim MazeTop As Double
    Dim NewWall As Shape

    MazeTop = Ws.Rows(MazeTopRow).Top
    
    For s = 1 To MazeSlots
        If s = 1 Or s = MazeSlots Then
            isWall(s) = True
        ElseIf ForceOpen Then
            isWall(s) = False
        Else
            isWall(s) = (Rnd < WallDensity)
        End If
    Next s
    
    If Not ForceOpen Then
        NewSlot = PathSlot + Int(Rnd * (2 * MaxTurn + 1)) - MaxTurn
        If NewSlot < 2 Then NewSlot = 2
        If NewSlot > MazeSlots - 1 Then NewSlot = MazeSlots - 1
        If NewSlot < PathSlot Then
            lo = NewSlot: hi = PathSlot
        Else
            lo = PathSlot: hi = NewSlot
        End If
        For s = lo To hi
            isWall(s) = False
        Next s
        
        PathSlot = NewSlot
    End If
    
    RunStart = 0
    For s = 1 To MazeSlots + 1
        If s <= MazeSlots Then IsW = isWall(s) Else IsW = False
        If IsW And RunStart = 0 Then RunStart = s
        If (Not IsW) And RunStart > 0 Then
            Set NewWall = Ws.Shapes.AddShape(msoShapeRectangle, x, MazeTop + (RunStart - 1) * BlockH, BlockW, (s - RunStart) * BlockH)
            NewWall.Fill.ForeColor.RGB = MazeColor
            NewWall.Line.Visible = msoFalse
            Blocks.Add NewWall
            RunStart = 0
        End If
    Next s
    
    Dim c As Shape
    Dim CoinSlot As Long
    Dim tries As Long
    
    ColCounter = ColCounter + 1
    If ColCounter >= CoinEveryCols And Not ForceOpen Then
        ColCounter = 0
        
        If CoinOnSafePath Then
            CoinSlot = PathSlot
        
        Else
            Do
                CoinSlot = Int(Rnd * (MazeSlots - 2)) + 2
                tries = tries + 1
            Loop While isWall(CoinSlot) And tries < 30
            If isWall(CoinSlot) Then CoinSlot = PathSlot
            
        End If
        
        Ws.Unprotect
        CoinShape.Visible = msoTrue
        CoinShape.Duplicate
        Ws.Protect UserInterfaceOnly:=True
        Set c = Ws.Shapes(Ws.Shapes.Count)
        If CoinSize > 0 Then
            c.LockAspectRatio = msoTrue
            c.Width = CoinSize
        End If
        c.Visible = msoTrue
        c.Left = x + (BlockW - c.Width) / 2
        c.Top = MazeTop + (CoinSlot - 1) * BlockH + (BlockH - c.Height) / 2
        Coins.Add c
    End If
        
        
End Sub

' --------------------------------------------------------------------------------------------- Game Functions
Function CheckBoxCollision(ByVal testLeft As Double, ByVal testTop As Double, ByVal pWidth As Double, ByVal pHeight As Double, ByVal Blocks As Collection) As Boolean
    Dim wall As Shape
    Dim wl As Double
    Dim collided As Boolean
    collided = False

    For Each wall In Blocks
        wl = wall.Left
        If testLeft + pWidth > wl Then
            If testLeft < wl + wall.Width Then
                If testTop + pHeight > wall.Top Then
                    If testTop < wall.Top + wall.Height Then
                        collided = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next wall

    CheckBoxCollision = collided
End Function

Function SlideDuckX(ByVal Player As Shape, ByVal dx As Double, ByVal offX As Double, ByVal offY As Double, ByVal w As Double, ByVal h As Double, ByVal Blocks As Collection) As Double
    Dim px As Double, py As Double
    Dim n As Long, k As Long, b As Long
    Dim sd As Double, lo As Double, hi As Double, md As Double

    px = Player.Left
    py = Player.Top

    If dx <> 0 Then
        n = Int(Abs(dx) / MaxStepPx) + 1
        sd = dx / n
        For k = 1 To n
            If CheckBoxCollision(px + sd + offX, py + offY, w, h, Blocks) Then
                lo = 0: hi = sd
                For b = 1 To 4
                    md = (lo + hi) / 2
                    If CheckBoxCollision(px + md + offX, py + offY, w, h, Blocks) Then hi = md Else lo = md
                Next b
                px = px + lo
                Exit For
            End If
            px = px + sd
        Next k
    End If

    SlideDuckX = px
End Function

Function BoxTouchesShape(ByVal s As Shape, ByVal bl As Double, ByVal bt As Double, ByVal bw As Double, ByVal bh As Double) As Boolean
    BoxTouchesShape = (bl + bw > s.Left) And (bl < s.Left + s.Width) And (bt + bh > s.Top) And (bt < s.Top + s.Height)
End Function

' --------------------------------------------------------------------------------------------- Main Program
Sub PlayGame()
    
    Dim StepK As Double
    Dim Coin As Shape
    Dim CurSpeed As Double
    Dim ColsSpawned As Long
    Dim Score As Double
    Dim BScore As Double
    Set Ws = ActiveSheet
    Call SetVariables
    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim FacingRight As Boolean
    Dim Blocks As Collection
    Dim wall As Shape
    Dim DisSinceSpawn As Double
    Dim MazePx As Double
    Dim DuckPx As Double
    Dim Spawned As Boolean
    Dim Pushed As Long
    Dim NextTop As Double
    Dim dx As Double
    Dim x As Double
    Dim hbW As Double
    Dim hbH As Double
    Dim hbOffX As Double
    Dim hbOffY As Double
    Dim relCells As Double
    
    Ws.Range("DO3:IB48").Interior.Color = PriColor
    Ws.Range("DO2:IB2").Interior.Color = SecColorB
    Controls.Fill.ForeColor.RGB = SecColorA
    Controls.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PriColor
    Decore.Fill.ForeColor.RGB = SecColorA
    Mes.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = SecColorB
    StartButton.Fill.ForeColor.RGB = SecColorB
    StartButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    Mes.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    ScoreBoard.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    BestScore.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BestColor
    Count.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ScoreColor
    
    Call StartEnd("Start")
    Set Blocks = New Collection
    Set Coins = New Collection
    
    BScore = Val(Mid(BestScore.TextFrame.Characters.Text, InStr(BestScore.TextFrame.Characters.Text, ":") + 1))
    FacingRight = True
    MazePx = ColLenght * MazeCellTick
    DuckPx = ColLenght * DuckCellTick
    hbW = Player.Width * HitboxScaleX
    hbH = Player.Height * HitboxScaleY
    
    Randomize
    
    PathSlot = Int((Player.Top + Player.Height / 2 - Ws.Rows(MazeTopRow).Top) / BlockH) + 1
    If PathSlot < 2 Then PathSlot = 2
    If PathSlot > MazeSlots - 1 Then PathSlot = MazeSlots - 1
    
    n = Int((SpawnPixelX - DeletePixelX - 0.001) / BlockW)
    For k = n To 0 Step -1
        x = SpawnPixelX - k * BlockW
        SpawnMazeColumn x, _
            (x + BlockW > Player.Left - ClearBehind * BlockW) And (x < Player.Left + Player.Width + ClearAhead * BlockW), _
            Blocks
    Next k
    
    DisSinceSpawn = 0
    Player.ZOrder msoBringToFront
    CurSpeed = StartGameSpeed
    ColCounter = 0
    GameRunning = True
    Application.ScreenUpdating = True
    hbOffX = (Player.Width - hbW) / 2
    hbOffY = (Player.Height - hbH) / 2
    
    
    Do While GameRunning ' -------------------------------------------------------------------- Game Loop
        
        StepK = CurSpeed * TickMs / BaseTickMs     ' scales every movement this tick
        
        If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then GameRunning = False
        ' Moves the maze and deleat blocks at the end
        For i = Blocks.Count To 1 Step -1
                Set wall = Blocks(i)
                wall.Left = wall.Left - MazePx * StepK
                If wall.Left <= DeletePixelX Then
                    wall.Delete
                    Blocks.Remove i
                    Score = Score + 1
                    ScoreBoard.TextFrame.Characters.Text = "Curent Score:" & Score
                End If
        Next i
        
        
        ' This insures that a new column is spawned only when a full one passed
        DisSinceSpawn = DisSinceSpawn + MazePx * StepK
        Spawned = False
        Do While DisSinceSpawn >= BlockW
            DisSinceSpawn = DisSinceSpawn - BlockW
            SpawnMazeColumn SpawnPixelX - DisSinceSpawn, False, Blocks
            Spawned = True
            ColsSpawned = ColsSpawned + 1
            If ColsSpawned Mod SpeedUpEveryCols = 0 Then
                CurSpeed = CurSpeed * SpeedMultiplier
                If CurSpeed > MaxGameSpeed Then CurSpeed = MaxGameSpeed
            End If
            
        Loop
        If Spawned Then Player.ZOrder msoBringToFront
        
        Pushed = 0
        Do While CheckBoxCollision(Player.Left + hbOffX, Player.Top + hbOffY, hbW, hbH, Blocks) And Pushed < MaxPushPx
            Player.Left = Player.Left - 1
            Pushed = Pushed + 1
        Loop
        
        Player.Left = SlideDuckX(Player, -DuckPx * StepK, hbOffX, hbOffY, hbW, hbH, Blocks)
        
        If Player.Left <= DangerPixelX Then GameRunning = False
        
        If GameRunning Then
            
            NextTop = Player.Top
            ' Vertical Movement (W)
            If GetAsyncKeyState(VK_W) <> 0 And GetAsyncKeyState(VK_S) = 0 Then
                NextTop = NextTop - RowLenght * StepK
                
            End If
            
            ' Vertical Movement (S)
            If GetAsyncKeyState(VK_S) <> 0 And GetAsyncKeyState(VK_W) = 0 Then
                NextTop = NextTop + RowLenght * StepK
                
            End If
            
            If NextTop <> Player.Top Then
                If Not CheckBoxCollision(Player.Left + hbOffX, NextTop + hbOffY, hbW, hbH, Blocks) Then
                    Player.Top = NextTop
                End If
            End If
            
            relCells = 0
            ' Horizontal Movement (A/D) & Flip Left
            If GetAsyncKeyState(VK_A) <> 0 And GetAsyncKeyState(VK_D) = 0 Then
                relCells = ATotal + DuckCellTick
                If FacingRight Then Player.Flip msoFlipHorizontal: FacingRight = False
            ElseIf GetAsyncKeyState(VK_D) <> 0 And GetAsyncKeyState(VK_A) = 0 Then
                relCells = DTotal + DuckCellTick
                If Not FacingRight Then Player.Flip msoFlipHorizontal: FacingRight = True
            End If
            
            If relCells <> 0 Then
                dx = relCells * ColLenght * StepK
                If dx > 0 Then                                            ' moving right: don't go past the wall
                    If dx > RightWallPixelX - (Player.Left + Player.Width) Then dx = RightWallPixelX - (Player.Left + Player.Width)
                    If dx < 0 Then dx = 0
                End If
                Player.Left = SlideDuckX(Player, dx, hbOffX, hbOffY, hbW, hbH, Blocks)
            End If
        End If
        
        
        
        For i = Coins.Count To 1 Step -1
            Set Coin = Coins(i)
            Coin.Left = Coin.Left - MazePx * StepK
            If Coin.Left <= DeletePixelX Then
                Coin.Delete
                Coins.Remove i
            ElseIf BoxTouchesShape(Coin, Player.Left + hbOffX, Player.Top + hbOffY, hbW, hbH) Then
                Coin.Delete
                Coins.Remove i
                Score = Score + CoinValue
                ScoreBoard.TextFrame.Characters.Text = "Curent Score:" & Score
            End If
        Next i
            
        Wait TickMs
        
    Loop
    
    For i = Blocks.Count To 1 Step -1
        Blocks(i).Delete
    Next i
    Set Blocks = Nothing
    For i = Coins.Count To 1 Step -1
        Coins(i).Delete
    Next i
    Set Coins = Nothing
    If Not FacingRight Then Player.Flip msoFlipHorizontal
    If Score > BScore Then BestScore.TextFrame.Characters.Text = "Best Score:" & Score
    Call StartEnd("End")
    
End Sub





