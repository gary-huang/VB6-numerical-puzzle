VERSION 5.00
Begin VB.Form frmFinalProject 
   AutoRedraw      =   -1  'True
   Caption         =   "Puzzle Board"
   ClientHeight    =   7605
   ClientLeft      =   3780
   ClientTop       =   3255
   ClientWidth     =   6450
   Icon            =   "Final_Project_HuangG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6450
   Begin VB.Frame fraNumericalBoard 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   2
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   3
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   5
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   6
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   7
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   8
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   9
         Left            =   840
         TabIndex        =   13
         Top             =   1440
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   10
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   11
         Left            =   2040
         TabIndex        =   11
         Top             =   1440
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   12
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   13
         Left            =   840
         TabIndex        =   9
         Top             =   2040
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   14
         Left            =   1440
         TabIndex        =   8
         Top             =   2040
         Width           =   600
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "1"
         Height          =   600
         Index           =   15
         Left            =   2040
         TabIndex        =   7
         Top             =   2040
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "End"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   5760
      Width           =   4695
      Begin VB.Label lblNumMoves 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Total Moves:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elapsed Time"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   480
      Top             =   6600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "Display High Scores"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuNumerical 
         Caption         =   "Numerical Board"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGraphical 
         Caption         =   "Graphical Board"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmFinalProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Final Project - Sliding Puzzle v1.0
'Author: Gary Huang
'Date: May 20th, 2014
'Files: Final_Project_HuangG.frm, Final_Project_HuangG.vbp,
'       Final_Project_HuangG.bas, Final_Project_HuangG.frx,
'       frmAbout.frm, frmAbout.frx, frmMsg.frm, frmMsg.frx
'Purpose: The purpose of this program is scramble a sliding
'         puzzle board, and allow the user to solve the
'         numerical board puzzle. The high score feature
'         and graphical board will be implemented in
'         version 2.0.

'Declare the form constant.

Const MAX = 4

'Declare the form variables.

Dim K As Integer, N As Integer
Dim NumMoves As Long
Dim Puzzle(1 To MAX, 1 To MAX) As PuzzleRec
Dim Diff As Date
Dim Start As Date

Option Explicit

Private Sub cmdExit_Click()
    
    End
    
End Sub

Private Sub cmdTile_Click(Index As Integer)
    
    Dim Direction As Integer
    Dim Msg As String
    
    'Determine if game has been won.
    
    If Not IsDonePuzzle(cmdTile()) Then
        
        'Determine the move direction.
        
        Direction = MoveDirection(Index, Puzzle())
        If Direction <> 0 Then
            
            'Move the tile if move is possible.
            
            MakeMove cmdTile(), Index, Direction, Puzzle()
            NumMoves = NumMoves + 1
            UpdateNumMoves lblNumMoves, NumMoves
            
            'Check if the puzzle is complete.
            
            If IsDonePuzzle(cmdTile()) Then
                
                'Stop the timer.
                
                tmrTime.Enabled = False
                
                'Display winning message.
                
                Msg = "You win! You took " & NumMoves & " moves"
                Msg = Msg & Chr$(10) & Chr$(13) & "Time taken: " & lblTime.Caption & "."
                MsgBox Msg, vbInformation, "Winning!"
                
                'Start a new numeric board game if user wishes to.
                
                If PlayAgain() Then
                    NewNumericGame NumMoves, Start, tmrTime, lblTime, lblNumMoves, cmdTile(), Puzzle()
                End If
            End If
        End If
    Else
        tmrTime.Enabled = False
        
        'Display winning message.
        
        Msg = "You won already! You took " & NumMoves & " moves."
        Msg = Msg & Chr$(10) & Chr$(13) & "Time taken: " & lblTime.Caption & "."
        MsgBox Msg, vbInformation, "Winning!"
        
        'Start new game if user wishes.
        
        If PlayAgain() Then
            NewNumericGame NumMoves, Start, tmrTime, lblTime, lblNumMoves, cmdTile(), Puzzle()
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
        
    'Initialize the tiles.
    
    For K = 0 To 15
        cmdTile(K).Caption = K + 1
    Next K
    
    Randomize
    
    'Start the timer.
    
    Start = Now
    tmrTime.Enabled = True

    'Scramble the numeric board.
    
    ScrambleNumericBoard cmdTile(), Puzzle()
    fraNumericalBoard.Visible = True
    MsgBox "Press Ok when ready.", vbInformation, "Ready?"
    
End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
    
    End
    
End Sub

Private Sub mnuGraphical_Click()
       
    ShowMsgBox 1, frmMsg
    
End Sub

Private Sub mnuHighScores_Click()
    
    ShowMsgBox 0, frmMsg
    
End Sub

Private Sub mnuNewGame_Click()
    
    NewNumericGame NumMoves, Start, tmrTime, lblTime, lblNumMoves, cmdTile(), Puzzle()
    
End Sub

Private Sub tmrTime_Timer()
    
    Dim Current As Date
    
    Current = Now
    
    'Calculate the difference between the current time and starting time to
    '   get the elapsed time.
    
    Diff = Current - Start
    lblTime.Caption = Format$(Diff, "hh:mm:ss")

End Sub
