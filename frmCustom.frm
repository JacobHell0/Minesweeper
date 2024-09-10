VERSION 5.00
Begin VB.Form frmCustom 
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHowManyBombs 
      Caption         =   "Play!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox txtGridX 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Text            =   "10"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.HScrollBar hsbGridX 
      Height          =   500
      Left            =   1320
      Max             =   50
      Min             =   2
      TabIndex        =   2
      Top             =   7440
      Value           =   10
      Width           =   8655
   End
   Begin VB.TextBox txtGridY 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "10"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.VScrollBar vsbGridY 
      Height          =   7815
      Left            =   720
      Max             =   50
      Min             =   2
      TabIndex        =   0
      Top             =   120
      Value           =   10
      Width           =   500
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdHowManyBombs_Click()
'Variable Table------------------------------------------------------------
Dim GridSize As Integer     'Stores the gridsize so it can be compared
'--------------------------------------------------------------------------
    'Makes sure the program cannot generate a grid with incorrect values
    If Val(txtGridY.Text) < 2 Or Val(txtGridX.Text) < 2 Then
        MsgBox ("Please enter a valid grid size")
    Else    'Asks the user how many bombs they want
        Do
            Do
                HowManyBombs = (InputBox("Please enter how many bombs you would like"))
            Loop Until IsNumeric(HowManyBombs) = True
            GridSize = (Val(txtGridY.Text) * Val(txtGridX.Text))
        Loop Until Val(HowManyBombs) > 0 And Val(HowManyBombs) < GridSize
        HowManyBombs = Int(Val(HowManyBombs))   'If the user entered a decimal, it strips off the decimal
        GridSizeX = Val(txtGridX.Text)
        GridSizeY = Val(txtGridY.Text)
        Unload Me
        frmMineSweeper.Visible = True
    End If
End Sub


Private Sub txtGridX_Change()   'The X value for the gridsize
    If Val(txtGridX.Text) < 51 Then
        If Val(txtGridX) > 2 Or Val(txtGridX) = 2 Then
            hsbGridX.Value = Round(Val(txtGridX.Text), 0)
        End If
    End If
End Sub
Private Sub hsbGridX_Change()
    txtGridX.Text = hsbGridX.Value
End Sub

Private Sub txtGridY_Change()   'The Y value for the gridsize
    If Val(txtGridY.Text) < 51 Then
        If Val(txtGridY) > 2 Or Val(txtGridY) = 2 Then
            vsbGridY.Value = Round(Val(txtGridY.Text), 0)
        End If
    End If
End Sub

Private Sub vsbGridY_Change()
    txtGridY.Text = vsbGridY.Value
End Sub
