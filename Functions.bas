Attribute VB_Name = "Functions"
'---------------------------------------------------------------------------------
'Variable Table
Global GridSizeX As Integer    'used to store the gridsizes size respectively
Global GridSizeY As Integer    'also needed for coordinate functions
Global HowManyBombs As String  'Used to store how many bombs the user inputs
Global Difficulty As Integer   'Used to store the difficulty when starting the game
Global c As Integer     'used to set the values of the for loop in scanning tiles and recursively clicking on the them
Global d As Integer     'also used to set values
Global e As Integer     'also used to set values
Global f As Integer     'also used to set values
Global FormUnloaded As Boolean  'Used to make sure no more recursive commands call formload after the form has been unloaded
Global Lose As Boolean  'Used to store whether or not to skip over the code if you lose also needs to be global so initialize can run
'----------------------------------------------------------------------------------------------------

Sub Main()
    frmMenu.Visible = True
End Sub

'Takes an index value and returns the x coordinate
Function CoordX(index As Integer) As Integer    'takes an index value and returns the X coordinate
    If index < GridSizeX + 1 Then
        CoordX = index
    ElseIf index Mod GridSizeX = 0 Then 'needed for if the tile is at the end
        CoordX = GridSizeX
    Else
        CoordX = index - GridSizeX * Int(index / GridSizeX)
    End If
End Function

'Takes an index value and returns the y coordinate
Function CoordY(index As Integer) As Integer    'takes an index value and returns the Y coordinate
    If index Mod GridSizeX = 0 Then     'needed for if the tile is at the end
        CoordY = Int(index / GridSizeX)
    Else
        CoordY = Int((index / GridSizeX)) + 1
    End If
End Function

'Takes the x and y coordinates and returns the index
Function Pos(x As Integer, Y As Integer) As Integer     'takes an x and y value and returns an index
    If Y = 0 Then   'needed if Y = 0 then the Pos will return a negative number
        Y = 1
    End If
    Pos = ((Y - 1) * GridSizeX) + x
End Function

'Determines if the tile being scanned is a corner, side, or a middle tile
Sub WhichTile(index As Integer, ByRef c As Integer, ByRef d As Integer, ByRef e As Integer, ByRef f As Integer)
    'Middle / sets initial values
    c = -1
    d = 1
    e = -1
    f = 1
    'Checks if it is on an edge/corner
    If CoordX(index) = 1 Then 'left
        e = e + 1
    End If
    If CoordX(index) = GridSizeX Then 'right
        f = f - 1
    End If
    If CoordY(index) = 1 Then ' top
        c = c + 1
    End If
    If CoordY(index) = GridSizeY Then ' bottom
        d = d - 1
    End If
End Sub
