Attribute VB_Name = "Module1"
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

'Declare the record type.

Type PuzzleRec
    Value As Integer
End Type

Option Explicit

'This general procedure scrambles the numeric puzzle board.

Sub ScrambleNumericBoard(Object As Object, Puzzle() As PuzzleRec)
        
    Dim K As Integer
    
    For K = 1 To 16
        Object(K - 1).Caption = K
    Next K
        
    RandomUniqueArray 0, 15, Object, 15
    MakeMemoryPuzzle Object, Puzzle()

End Sub

'This general procedure creates a copy of the puzzle in memory.

Sub MakeMemoryPuzzle(Object As Variant, Puzzle() As PuzzleRec)
    
    Dim X As Integer, Y As Integer, K As Integer
    
    For K = 0 To 15
        X = K Mod 4 + 1
        Y = K \ 4 + 1
        Puzzle(X, Y).Value = Val(Object(K).Caption)
        If Object(K).Caption = "16" Then
            Object(K).Visible = False
        Else
            Object(K).Visible = True
        End If
    Next K
    
End Sub

'This general procedure generates a random array of integers by swapping positions of already existent values.

Sub RandomUniqueArray(Low As Integer, High As Integer, Values As Variant, MaxControlIndex As Integer)
        
    Dim K As Integer, N As Integer
    Dim SwapPos As Integer, Temp As Integer
    
    For K = 0 To 15
        SwapPos = MakeRandom(Low, High)
        Temp = Values(SwapPos).Caption
        Values(SwapPos).Caption = Values(K).Caption
        Values(K).Caption = Temp
    Next K
        
End Sub

'This function makes a random number in between the given range.

Function MakeRandom(Low As Integer, High As Integer) As Integer

    Dim Temp As Integer
    
    Temp = Int(Rnd * (High - Low + 1)) + Low
    
    MakeRandom = Temp
    
End Function

'This general procedure updates the number of moves.

Sub UpdateNumMoves(LabelName As Label, ByVal NumMoves As Long)
    
    LabelName.Caption = NumMoves
    
End Sub

'This general procedure makes a move on the memory puzzle and visual puzzle.

Sub MakeMove(Object As Variant, Index As Integer, Direction As Integer, Puzzle() As PuzzleRec)
    
    Dim CurX As Integer, CurY As Integer
    Dim TargetX As Integer, TargetY As Integer
    Dim Diff As Integer
    
    CurX = Index Mod 4 + 1
    CurY = Index \ 4 + 1
    Select Case Direction
        Case 12
            TargetX = CurX
            TargetY = CurY - 1
            Diff = -4
        Case 3
            TargetX = CurX + 1
            TargetY = CurY
            Diff = 1
        Case 6
            TargetX = CurX
            TargetY = CurY + 1
            Diff = 4
        Case 9
            TargetX = CurX - 1
            TargetY = CurY
            Diff = -1
    End Select
    
    Swap Puzzle(CurX, CurY), Puzzle(TargetX, TargetY)
    Object(Index).Visible = False
    Object(Index).Caption = Puzzle(CurX, CurY).Value
    Object(Index + Diff).Visible = True
    Object(Index + Diff).Caption = Puzzle(TargetX, TargetY).Value
    
End Sub

'This general procedure swaps the value of 2 puzzle pieces.

Sub Swap(Val1 As PuzzleRec, Val2 As PuzzleRec)
    
    Dim Temp As Integer
    
    Temp = Val1.Value
    Val1.Value = Val2.Value
    Val2.Value = Temp
    
End Sub

'This function determines the move direction of the clicked piece.

Function MoveDirection(ClickedIndex As Integer, Puzzle() As PuzzleRec) As Integer
    
    Dim CurX As Integer, CurY As Integer
    Dim TargetX As Integer, TargetY As Integer
    Dim K As Integer, N As Integer
    Dim DiffX As Integer, DiffY As Integer
    Dim Temp As Integer
    
    CurY = ClickedIndex \ 4 + 1
    CurX = ClickedIndex Mod 4 + 1
    
    For K = 1 To 4
        For N = 1 To 4
            If Puzzle(K, N).Value = 16 Then
                TargetX = K
                TargetY = N
            End If
        Next N
    Next K
    
    DiffX = CurX - TargetX
    DiffY = CurY - TargetY
    
    If Abs(DiffX) > 1 Or Abs(DiffY) > 1 Then
        Temp = 0
    ElseIf Abs(DiffX) = 1 And Abs(DiffY) = 1 Then
        Temp = 0
    ElseIf Abs(DiffX) = 1 And Abs(DiffY) = 0 Then
        If DiffX = 1 Then
            Temp = 9
        Else
            Temp = 3
        End If
    ElseIf Abs(DiffX) = 0 And Abs(DiffY) = 1 Then
        If DiffY = 1 Then
            Temp = 12
        Else
            Temp = 6
        End If
    End If
    
    MoveDirection = Temp
    
End Function

'This function checks whether the puzzle is complete.

Function IsDonePuzzle(Object As Variant) As Boolean
    
    Dim K As Integer
    Dim Valid As Boolean
        
    Valid = True
        
    For K = 0 To 15
        If Object(K).Caption <> K + 1 Then
            Valid = False
        End If
    Next K
    
    IsDonePuzzle = Valid
    
End Function

'This function determines whether the user wishes to play again.

Function PlayAgain() As Boolean

    Dim Response As Integer
    Dim DType As Integer
    
    DType = vbYesNo + vbQuestion
    
    Response = MsgBox("Would you like to play again?", DType, "Play Again?")
    
    If Response = vbYes Then
        PlayAgain = True
    Else
        PlayAgain = False
    End If

End Function

'This general procedure starts a new numeric board game.

Sub NewNumericGame(NumMoves As Long, Start As Date, TimerName As Timer, ElapsedTime As Label, NumMovesLabel As Label, CommandBttnArray As Object, Puzzle() As PuzzleRec)
    
    MsgBox "Press Ok when ready.", vbInformation, "Ready?"
    
    Start = Now
    
    ElapsedTime.Caption = "00:00:00"
    
    TimerName.Enabled = True
    
    ScrambleNumericBoard CommandBttnArray, Puzzle()
    
    NumMoves = 0
    UpdateNumMoves NumMovesLabel, NumMoves
    
End Sub

'This general procedure shows a message to the user.

Sub ShowMsgBox(Mode As Integer, FormName As Form)
    
    Const Msg0 = "The high score feature will be implemented in version 2.0."
    Const Msg1 = "The graphical version of this game will be implemented in version 2.0."
      
    Select Case Mode
        Case 0, 1
            FormName.Caption = "Coming Soon"
    End Select
    
    Select Case Mode
        Case 0
            FormName!lblMsg.Caption = Msg0
        Case 1
            FormName!lblMsg.Caption = Msg1
    End Select
    
    FormName.Show vbModal
    
End Sub
