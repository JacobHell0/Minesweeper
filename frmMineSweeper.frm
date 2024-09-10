VERSION 5.00
Begin VB.Form frmMineSweeper 
   Caption         =   "MineSweeper"
   ClientHeight    =   10395
   ClientLeft      =   4785
   ClientTop       =   1395
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   10995
   Begin VB.CommandButton cmdLatch1Control 
      Caption         =   "Recursive Delay"
      Height          =   615
      Left            =   9960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer tmrLatch2 
      Interval        =   100
      Left            =   1080
      Top             =   9840
   End
   Begin VB.Timer tmrLatch 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   600
      Top             =   9840
   End
   Begin VB.PictureBox PicBackDrop 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   960
      Picture         =   "frmMineSweeper.frx":0000
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   588
      TabIndex        =   1
      Top             =   120
      Width           =   8820
      Begin VB.Label lblTimer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   7680
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblBombCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   9840
   End
   Begin VB.CommandButton cmdTile 
      DisabledPicture =   "frmMineSweeper.frx":6952
      Height          =   735
      Index           =   0
      Left            =   240
      Picture         =   "frmMineSweeper.frx":7788
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMineSweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------
'                Recursive Minesweeper
'                By: Jacob Rempel
'                ICS4U1
'                April 9, 2021
'----------------------------------------------------
'Purpose
'   The purpose of this code is to replicate the windows 98 minesweeper with some creative liberties. As well
'as accurately utilize the powerful programming tool recursion.

'----------------------------------------------------
'Design Decisions
'-I decided to have a menu/splashscreen instead of a dropdown menu along the top
'-I also decided to have a cap on the maximum x and y values when you are setting your grid, because anything over 42x43 seems to overflow when you click on the grid with 1 bomb
'-----------------------------------------------------
Option Explicit

'Variable Table------------------------------------------------------------
Dim Flag As Boolean         'Tests if it is the first click or not
Dim BombCounter As Integer 'delete (this is for debugging purposes)
Dim FlagCounter As Integer 'needed to display how many flags are placed
Dim BombsGenerated As Boolean 'needed so when you right click before the bombs generate it does not screw up the win condition
Dim GridSize As Integer       'Stores the grid size
Dim ClickCounter As Integer     'Stores the number of times the user triggered a tile for the win condition
Dim Clock As Integer            'Used to store the integer for the display (timer)
Dim Wait As Boolean         'Used to latch the click event
Dim Wait2 As Boolean        'Used to latch the bombs being shown when you lose the game

'Constant Table-------------------
Const Red As Double = &H8080FF          'Makes some items more readable
Const White As Double = &H8000000F      'in the right click (flag) functions
'--------------------------------------


Private Sub cmdLatch1Control_Click()    'Turns the latch on or off
    If tmrLatch.Enabled = False Then
        tmrLatch.Enabled = True
    Else
        tmrLatch.Enabled = False
    End If
End Sub

Private Sub cmdTile_Click(index As Integer)
'Variable Table----------------------------
Dim a As Integer    'Controls the X value of the for loop that scans the tiles
Dim b As Integer    'Controls the Y value of the for loop that scans the tiles
Dim Counter As Integer  'temporary counter for holding the .tag value generated by scanning the tiles around where the user clicked
'------------------------------------------

    If Lose = True Then     'Lose will never = true until LoseGame or WinGame is called so if Lose = true then
        Call Initialize     'we know that it is the first time being run
    End If
    If FormUnloaded = False Then
        Wait = True 'Latch for slight delay
        Do While Wait = True And tmrLatch.Enabled = True
            DoEvents
        Loop
        If cmdTile(index).BackColor = White Then    'Make sure the tile is not clicked on if it is a flag
            
            'Generates bombs only on the first click
            If Flag = False Then
                Call Bombs(Val(HowManyBombs), index)
                Flag = True
            End If
    
            'controls if you click on a bomb
            If cmdTile(index).Tag = "9" Then
                Call WinAndLose(index, False)   'Calls the lose game function
            End If
            
            If Lose = False Or FormUnloaded = False Then  'Makes sure the rest of the code is not run if you lose/win or if the form is unloaded
                Call WhichTile(index, c, d, e, f)   'Figures out what values c,d,e,f should before the loop
                'Scanning Tiles-----------------------------------------------------------------
                Counter = 0     'Needed so if their are no bombs around the tile, it does not set the .tag to null
                For a = c To d
                    For b = e To f
                        If cmdTile(Pos(CoordX(index) + b, CoordY(index) + a)).Tag = "9" Then
                            Counter = Counter + 1
                        End If
                    Next b                      'Cannot combine the recursion loop and the scanning
                Next a                          'tiles loop because the recursive loop needs the .tag value
                cmdTile(index).Tag = Counter    'However, we can still use the same c,d,e,f values
                Call Graphics(index) 'updates the tiles graphics
                'Recursive bit---------------------------------------------------------------
                For a = c To d
                    For b = e To f
                        If FormUnloaded = False Then    'Used so when the form is unloaded, id doesn't call form load when finishing up the recursion
                            If cmdTile(index).Tag = 0 And Pos(CoordX(index) + b, CoordY(index) + a) < GridSize + 1 Then     'makes sure it does not click outside the grid
                                If cmdTile(Pos(CoordX(index) + b, CoordY(index) + a)).Tag = "" And cmdTile(Pos(CoordX(index) + b, CoordY(index) + a)).BackColor = White Then
                                    cmdTile_Click (Pos(CoordX(index) + b, CoordY(index) + a))
                                End If
                            End If
                        End If
                    Next b
                Next a
                If FormUnloaded = False Then    'Prevents WinGame from being called if a recursive spawn is resolving
                    ClickCounter = ClickCounter + 1 'Checks if you have won by clicking on every tile
                    If GridSize - ClickCounter = HowManyBombs Then
                        Call WinAndLose(index, True)    'Calls the win game function
                    End If
                End If
            End If  'lose or form unloaded end if
        End If  'flag backcolour end if
    End If  'Form Unloaded end if
End Sub

Sub Graphics(index As Integer)  'only called for bombs and numbers
    cmdTile(index).Enabled = False  'Disable buttons
    cmdTile(index).DisabledPicture = LoadPicture(".\Graphics\" & cmdTile(index).Tag & ".bmp", 4, 0, cmdTile(index).Width, cmdTile(index).Height)
End Sub

Sub WinAndLose(index As Integer, Win As Boolean)
'Variable Table----------------------------
Dim a As Integer    'Used in the for loop that goes the the grid and reveals the bombs
Dim Endgame As Integer  'Used to temporarily store the result of the user's yes or no
Dim Img As String       'Used to either display the bombs or the defused bombs
'------------------------------------------
    
    If Win = True Then      'If you win
        MsgBox ("You Win!")
        Img = "Defused"     'Used later to set the graphics accordingly
    ElseIf Win = False Then 'If you lose
        cmdTile(index).Picture = LoadPicture(".\Graphics\mine.bmp") 'reveals the mine you clicked on
        MsgBox ("You Blew Up")
        Img = ""   'Sets Img to no string so, later, the for loop uses the picture "mine.bmp"
    End If
    
    For a = 1 To GridSize   'Reveals all the bombs to the user
        If cmdTile(a).BackColor = Red And cmdTile(a).Tag <> "9" Then    'will never happen on win
            Wait2 = True    'Latch to make the reveal slower (added because visually appealing)
            Do While Wait2 = True
                DoEvents
            Loop
            cmdTile(a).Picture = LoadPicture(".\Graphics\mineX.bmp")    'Reveals if the user incorrectly tagged a bomb
        End If
        
        If cmdTile(a).Tag = "9" Then
            Wait2 = True    'Had to use 2 latches because if the latch is outside this if statement, it latches tiles with no bombs/flags on them
            Do While Wait2 = True
                DoEvents
            Loop
            cmdTile(a).Picture = LoadPicture(".\Graphics\mine" & Img & ".bmp")
        End If
    Next a
    
    Endgame = MsgBox("Would you like to play again?", vbYesNo, "MineSweeper")
    If Endgame = 6 Then 'If user clicks yes
        frmMenu.Visible = True  'Unloads form and goes to menu
        Unload Me
        FormUnloaded = True
        Lose = True
    Else    'If user clicks no
        End
    End If
End Sub
Sub Display()   'Used to update the flag counter at the top of the screen
    Dim Bombs As Integer
    Bombs = Val(HowManyBombs)
    If Bombs - FlagCounter > 9 Then
        lblBombCounter.Caption = Bombs - FlagCounter
    ElseIf Bombs - FlagCounter > -1 Then
        lblBombCounter.Caption = "0" & Bombs - FlagCounter
    ElseIf Bombs - FlagCounter < -9 Then
        lblBombCounter.Caption = Bombs - FlagCounter
    ElseIf Bombs - FlagCounter < 0 Then
        lblBombCounter.Caption = "-0" & Abs(Bombs - FlagCounter)
    End If
End Sub


Private Sub cmdTile_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Right click to place a flag/marker function
    If Button = 2 And BombsGenerated = True Then    'If Right Click, BombsGenerated = true prevents the user from right click when no bombs have been generated
        If cmdTile(index).BackColor = White Then    'If the .backcolor is white then make it red
            cmdTile(index).BackColor = Red
            cmdTile(index).Picture = LoadPicture(".\Graphics\Minesweeper_flag.bmp")
            FlagCounter = FlagCounter + 1
            Call Display                        'Updates (the flag counter of the) Display
            If cmdTile(index).Tag = "9" Then
                BombCounter = BombCounter + 1
            Else
                BombCounter = BombCounter - 1
            End If
        Else    'If red
            cmdTile(index).BackColor = White    'original white colour
            cmdTile(index).Picture = LoadPicture(".\Graphics\Default.bmp")
            FlagCounter = FlagCounter - 1
            Call Display        'Updates display
            If cmdTile(index).Tag = "9" Then        'Bomb counter / win condition
                BombCounter = BombCounter - 1
            Else
                BombCounter = BombCounter + 1
            End If
        End If
        If BombCounter = Val(HowManyBombs) Then     'Win Condition (checks if you have flagged all the bombs)
            tmrClock.Enabled = False
            Call WinAndLose(index, True)    'Calls the win game function
        End If
    End If
End Sub

Private Sub Initialize()    'Initializes values after a game has been won/lost
    Flag = False
    Lose = False
    BombsGenerated = False
    ClickCounter = False
    FormUnloaded = False
End Sub


Sub GenGrid()   'Generates the grid
'Variable Table----------------------------
Dim a As Integer
'------------------------------------------
    GridSize = GridSizeX * GridSizeY
    frmMineSweeper.Visible = True
    For a = 1 To GridSize
        Load cmdTile(a)
        With cmdTile(a)
            .Visible = True
            If (a - 1) Mod GridSizeX = 0 And (a - 1) <> 0 Then
                .Left = cmdTile(a - GridSizeX).Left
            Else
                .Left = cmdTile(a - 1).Left + cmdTile(a - 1).Width
            End If
            If a > GridSizeX Then
                .Top = cmdTile(a - GridSizeX).Top + cmdTile(a - GridSizeX).Height
            End If
        End With
    Next a
End Sub

Sub Bombs(Amount As Integer, index As Integer)
'Variable Table----------------------------
Dim x As Integer                'Used in the for loop
Dim RandomNumber As Integer     'Stores the random number that is generated by the RND function
Dim Counter As Integer          'Stores the number of bombs that have been successfully placed
'------------------------------------------
    Randomize   'Makes sure the bombs are random each time
        For x = 1 To Amount
            RandomNumber = Round((GridSize * Rnd), 0)
            If cmdTile(RandomNumber).Tag = "" And index <> RandomNumber Then 'makes your first click not a bomb
                cmdTile(RandomNumber).Tag = "9"
                Counter = Counter + 1
            End If
            If x = Amount And Counter <> Amount Then
                x = x - 1
            End If
        Next x
    BombsGenerated = True
End Sub

Private Sub Form_Load()
    'Display Variables need to be set before display updates or else initialize does not update them until the click event
    FlagCounter = 0
    BombCounter = 0
    Clock = 0
    Call StartGame
End Sub

Sub StartGame()
    'Presets for gridsizes and bomb amounts
    If Difficulty = 0 Then 'Easy
        GridSizeX = 10
        GridSizeY = 10
        HowManyBombs = Str(5)
    ElseIf Difficulty = 1 Then 'Medium
        GridSizeX = 10
        GridSizeY = 10
        HowManyBombs = Str(10)
    ElseIf Difficulty = 2 Then 'Hard
        GridSizeX = 15
        GridSizeY = 15
        HowManyBombs = Str(30)  'Need to use string values because HowManyBombs is a string
    End If
    
    Call GenGrid  'Generates the grid
    
    Call Display  'Sets the flag counter at the top of the screen
            
    'Correctly aligns the size of the form with the amount of tiles. However, these values are for my computer's screen size
    If cmdTile(GridSizeX).Left < 11220 Then
        frmMineSweeper.Width = 11220
    Else
        frmMineSweeper.Width = cmdTile(GridSizeX).Left + (cmdTile(GridSizeX).Width * 3)
    End If
    
    'Makes sure the values do not overflow
    If GridSizeX > 19 Then      'makes the window go to the top of your screen for more space
        frmMineSweeper.Left = 0 'if more space is needed
    End If
    If GridSizeY > 15 Then
        frmMineSweeper.Top = 0
    End If
    
    If GridSizeY > 20 Then  '20 is the max tiles I can fit on my screen
        frmMineSweeper.WindowState = 2  'this is to prevent overflow errors
    Else
        frmMineSweeper.Height = cmdTile(GridSizeY).Top + (1100 * GridSizeY)
    End If
    
    'Starts the clock that displays how long it took you to sweep the mines
    tmrClock.Enabled = True
End Sub

Private Sub tmrClock_Timer()    'Used for caption manipulation of the timer at the top of the screen
    Clock = Clock + 1   'The timer's interval is a second long
    If Clock < 10 Then
        lblTimer.Caption = "00" & Clock
    ElseIf Clock < 100 Then
        lblTimer.Caption = "0" & Clock
    ElseIf Clock < 1000 Then
        lblTimer.Caption = Clock
    Else
        lblTimer.Caption = "999"
    End If
End Sub

Private Sub tmrLatch_Timer()    'Used to control the speed of the click event
    Wait = False
End Sub

Private Sub tmrLatch2_Timer()   'Used to control the speed of the bombs being revealed when you lose
    Wait2 = False
End Sub
