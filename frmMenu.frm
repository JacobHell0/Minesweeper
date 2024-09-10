VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000007&
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHard 
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdMedium 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdEasy 
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   0
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Jacob Rempel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Image imgText 
      Height          =   5295
      Left            =   240
      Picture         =   "frmMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEasy_Click()
    Difficulty = 0
    frmMenu.Visible = False
    frmMineSweeper.Visible = True
End Sub

Private Sub cmdMedium_Click()
    Difficulty = 1
    frmMenu.Visible = False
    frmMineSweeper.Visible = True
End Sub

Private Sub cmdHard_Click()
    Difficulty = 2
    frmMenu.Visible = False
    frmMineSweeper.Visible = True
End Sub

Private Sub cmdCustom_Click()
    Difficulty = 3
    frmMenu.Visible = False
    frmCustom.Visible = True
    FormUnloaded = False
End Sub


