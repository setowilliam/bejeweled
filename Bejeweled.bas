Attribute VB_Name = "Bejeweled"
'Title: Bejeweled
'Author: William Seto
'Date: June 5, 2007
'Files: Bejeweled.bas, Bejeweled.frm, Bejeweled.frx, Bejeweled.vbp, Bejeweled.vbw, Bejeweled High Scores.txt,
'       frmHighScore.frm
'Purpose: The purpose of this program is to allow the user to play a replica of
'         Bejeweled.

Option Explicit

'This general procedure sets everything up to start a game.

Sub StartGame(Gem() As Integer, InitialGem As Integer, IsMove As Boolean)

    Dim k As Integer
    Dim IsValid As Integer
    
    InitialGem = -9999
    
    Do
        FillGemArray Gem()
        IsValid = ValidCheck(Gem())
        If IsValid = 0 Then
            IsMove = CheckForMoves(Gem())
        End If
    Loop While IsMove = False Or IsValid > 0
    
    For k = 0 To 63
        frmBejeweled!picGem(k).Picture = frmBejeweled!imgGemImage(Gem(k))
        frmBejeweled!picGem(k).Enabled = True
    Next k
    
End Sub

'This general procedure gets the high scores and their respective names.

Sub FillHighScores(Names() As String, Scores() As Long)

    Dim k As Integer
    
    k = 0
    
    frmHighScore.Print Tab(5); "Name"; Tab(31); "Score"
    frmHighScore.Print
    
    Open App.Path & "\Bejeweled High Scores.txt" For Input As #1
    
        Do While Not EOF(1)
            k = k + 1
            Input #1, Names(k)
            Input #1, Scores(k)
            frmHighScore.Print k; Tab(5); Names(k); Tab(30); Scores(k)
        Loop
    
    Close #1
    
End Sub

'This function determines if there are any moves left.

Function CheckForMoves(Gem() As Integer) As Boolean

    Dim k As Integer
    Dim Moves As Integer
    Dim Temp As Boolean
    
    Temp = False
    k = 0
    Moves = 0
    
    Do
        Select Case k
            Case 0 To 6, 8 To 14, 16 To 22, 24 To 30, 32 To 38, 40 To 46
                Moves = RightParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 1 To 7, 9 To 15, 17 To 23, 25 To 31, 33 To 39, 41 To 47
                Moves = LeftParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 0 To 39
                Moves = DownLinearCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 0 To 4, 8 To 12, 16 To 20, 24 To 28, 32 To 36, 40 To 44, 48 To 52, 56 To 60
                Moves = RightLinearCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 3 To 7, 11 To 15, 19 To 23, 27 To 31, 35 To 39, 43 To 47, 51 To 55, 59 To 63
                Moves = LeftLinearCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 24 To 63
                Moves = UpLinearCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 16 To 22, 24 To 30, 32 To 38, 40 To 46, 48 To 54, 56 To 62
                Moves = UpRightParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 17 To 23, 25 To 31, 33 To 39, 41 To 47, 49 To 55, 57 To 63
                Moves = UpLeftParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 8 To 13, 16 To 21, 24 To 29, 32 To 37, 40 To 45, 48 To 53, 56 To 61
                Moves = RightUpParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 0 To 5, 8 To 13, 16 To 21, 24 To 29, 32 To 37, 40 To 45, 48 To 53
                Moves = RightDownParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 10 To 15, 18 To 23, 26 To 31, 34 To 39, 42 To 47, 50 To 55, 58 To 63
                Moves = LeftUpParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 2 To 7, 10 To 15, 18 To 23, 26 To 31, 34 To 39, 42 To 47, 50 To 55
                Moves = LeftDownParallelCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 1 To 6, 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54
                Moves = DownTriangleCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54, 57 To 62
                Moves = UpTriangleCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 8 To 14, 16 To 22, 24 To 30, 32 To 38, 40 To 46, 48 To 54
                Moves = RightTriangleCheck(Gem(), k, Moves)
        End Select
        
        Select Case k
            Case 9 To 15, 17 To 23, 25 To 31, 33 To 39, 41 To 47, 49 To 55
                Moves = LeftTriangleCheck(Gem(), k, Moves)
        End Select
        
        k = k + 1
        
    Loop While Moves = 0 And k <= 63
    
    If Moves >= 1 Then
        Temp = True
    End If
    
    CheckForMoves = Temp
    
End Function

'This function determines whether there are any gems that are already connected in rows
'or columns of 3 or more.

Function ValidCheck(Gem() As Integer) As Integer

    Dim Valid As Integer
    Dim k As Integer
    
    k = -1
    Valid = 0
    
    Do
        k = k + 1
        Select Case k
            Case 0 To 5, 8 To 13, 16 To 21, 24 To 29, 32 To 37, 40 To 45, 48 To 53, 56 To 61
                Valid = Valid + RightRowCheck(Gem(), k)
        End Select
        
        Select Case k
            Case 2 To 7, 9 To 15, 17 To 23, 25 To 31, 33 To 39, 41 To 47, 49 To 55, 57 To 63
                Valid = Valid + LeftRowCheck(Gem(), k)
        End Select
        
        Select Case k
            Case 1 To 6, 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54, 57 To 62
                Valid = Valid + CenterRowCheck(Gem(), k)
        End Select
        
        Select Case k
            Case 0 To 47
                Valid = Valid + DownColumnCheck(Gem(), k)
        End Select
        
        Select Case k
            Case 16 To 63
                Valid = Valid + UpColumnCheck(Gem(), k)
        End Select
        
        Select Case k
            Case 8 To 55
                Valid = Valid + CenterColumnCheck(Gem(), k)
        End Select
    
    Loop While Valid = 0 And k <= 62
    
    ValidCheck = Valid
    
End Function

'This function determines if a move is illegal or not.

Function ValidMoveCheck(InitialGem As Integer, SecondaryGem As Integer) As Boolean

    Dim Temp As Boolean
    
    Temp = False
    
    Select Case InitialGem
        Case 0
            If SecondaryGem = 1 Or SecondaryGem = 8 Then
                Temp = True
            End If
        Case 7
            If SecondaryGem = 6 Or SecondaryGem = 15 Then
                Temp = True
            End If
        Case 56
            If SecondaryGem = 48 Or SecondaryGem = 57 Then
                Temp = True
            End If
        Case 63
            If SecondaryGem = 55 Or SecondaryGem = 62 Then
                Temp = True
            End If
        Case 1 To 6
            If SecondaryGem = InitialGem - 1 Or SecondaryGem = InitialGem + 1 Or SecondaryGem = InitialGem + 8 Then
                Temp = True
            End If
        Case 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54
            If SecondaryGem = InitialGem - 1 Or SecondaryGem = InitialGem + 1 Or SecondaryGem = InitialGem + 8 Or SecondaryGem = InitialGem - 8 Then
                Temp = True
            End If
        Case 57 To 62
            If SecondaryGem = InitialGem - 1 Or SecondaryGem = InitialGem + 1 Or SecondaryGem = InitialGem - 8 Then
                Temp = True
            End If
        Case 8, 16, 24, 32, 40, 48
            If SecondaryGem = InitialGem - 8 Or SecondaryGem = InitialGem + 8 Or SecondaryGem = InitialGem + 1 Then
                Temp = True
            End If
        Case 15, 23, 31, 39, 47, 55
            If SecondaryGem = InitialGem - 8 Or SecondaryGem = InitialGem + 8 Or SecondaryGem = InitialGem + -1 Then
                Temp = True
            End If
    End Select
    
    ValidMoveCheck = Temp
    
End Function

'This function checks whether there is a move in the specific pattern.

Function RightParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 9) And Gem(k) = Gem(k + 17) Then
        Moves = Moves + 1
    End If
    
    RightParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function LeftParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 7) And Gem(k) = Gem(k + 15) Then
        Moves = Moves + 1
    End If
    
    LeftParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function DownLinearCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 16) And Gem(k) = Gem(k + 24) Then
        Moves = Moves + 1
    End If
    
    DownLinearCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function RightLinearCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 2) And Gem(k) = Gem(k + 3) Then
        Moves = Moves + 1
    End If
    
    RightLinearCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function LeftLinearCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 2) And Gem(k) = Gem(k - 3) Then
        Moves = Moves + 1
    End If
    
    LeftLinearCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function UpLinearCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 16) And Gem(k) = Gem(k - 24) Then
        Moves = Moves + 1
    End If
    
    UpLinearCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function UpRightParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 7) And Gem(k) = Gem(k - 15) Then
        Moves = Moves + 1
    End If
    
    UpRightParallelCheck = Moves
        
End Function

'This function checks whether there is a move in the specific pattern.

Function UpLeftParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 9) And Gem(k) = Gem(k - 17) Then
        Moves = Moves + 1
    End If
    
    UpLeftParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function RightUpParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 6) And Gem(k) = Gem(k - 7) Then
        Moves = Moves + 1
    End If
    
    RightUpParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function RightDownParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 9) And Gem(k) = Gem(k + 10) Then
        Moves = Moves + 1
    End If
    
    RightDownParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function LeftUpParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 9) And Gem(k) = Gem(k - 10) Then
        Moves = Moves + 1
    End If
    
    LeftUpParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function LeftDownParallelCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 6) And Gem(k) = Gem(k + 7) Then
        Moves = Moves + 1
    End If
    
    LeftDownParallelCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function DownTriangleCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k + 7) And Gem(k) = Gem(k + 9) Then
        Moves = Moves + 1
    End If
    
    DownTriangleCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function UpTriangleCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 7) And Gem(k) = Gem(k - 9) Then
        Moves = Moves + 1
    End If
    
    UpTriangleCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function RightTriangleCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 7) And Gem(k) = Gem(k + 9) Then
        Moves = Moves + 1
    End If
    
    RightTriangleCheck = Moves
    
End Function

'This function checks whether there is a move in the specific pattern.

Function LeftTriangleCheck(Gem() As Integer, k As Integer, Moves As Integer) As Integer

    If Gem(k) = Gem(k - 9) And Gem(k) = Gem(k + 7) Then
        Moves = Moves + 1
    End If
    
    LeftTriangleCheck = Moves
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function RightRowCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k + 1) And Gem(k) = Gem(k + 2) Then
        Temp = Temp + 1
    End If
    
    RightRowCheck = Temp
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function LeftRowCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k - 1) And Gem(k) = Gem(k - 2) Then
        Temp = Temp + 1
    End If
    
    LeftRowCheck = Temp
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function CenterRowCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k - 1) And Gem(k) = Gem(k + 1) Then
        Temp = Temp + 1
    End If
    
    CenterRowCheck = Temp
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function DownColumnCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k + 8) And Gem(k) = Gem(k + 16) Then
        Temp = Temp + 1
    End If
    
    DownColumnCheck = Temp
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function UpColumnCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k - 8) And Gem(k) = Gem(k - 16) Then
        Temp = Temp + 1
    End If
    
    UpColumnCheck = Temp
    
End Function

'This function determines where there are 3 or more gems connected in the specified
'direction.

Function CenterColumnCheck(Gem() As Integer, k As Integer) As Integer

    Dim Temp As Integer
    
    If Gem(k) = Gem(k + 8) And Gem(k) = Gem(k - 8) Then
        Temp = Temp + 1
    End If
    
    CenterColumnCheck = Temp
    
End Function

'This general procedure fills the "Gem" array with random numbers that correspond to a
'certain gem.

Sub FillGemArray(Gems() As Integer)

    Dim k As Integer
    Dim GemNum As Integer
    
    For k = 0 To 63
        GemNum = Int(Rnd * 7) + 1
        Select Case GemNum
            Case 1
                Gems(k) = 0
            Case 2
                Gems(k) = 1
            Case 3
                Gems(k) = 2
            Case 4
                Gems(k) = 3
            Case 5
                Gems(k) = 4
            Case 6
                Gems(k) = 5
            Case 7
                Gems(k) = 6
        End Select
    Next k
    
End Sub

'This general procedure swaps places with 2 selected gems.

Sub SwapGem(FirstGem As Integer, SecondGem As Integer, Gems() As Integer)

    Dim Temp As Integer
    
    Temp = Gems(FirstGem)
    Gems(FirstGem) = Gems(SecondGem)
    Gems(SecondGem) = Temp
    frmBejeweled!picGem(FirstGem).Picture = frmBejeweled!imgGemImage(Gems(FirstGem))
    frmBejeweled!picGem(SecondGem).Picture = frmBejeweled!imgGemImage(Gems(SecondGem))
    
End Sub

'This general procedure destroys any gems that have 3 or more gems connected to each other,
'then drops the gems at the top down to fill in the empty spaces.

Sub DestroyGems(InitialGem As Integer, SecondaryGem As Integer, Gems() As Integer, HValidSwapOne As Boolean, HValidSwapTwo As Boolean, VValidSwapOne As Boolean, VValidSwapTwo As Boolean, CurrentScore As Long)

    Dim k As Integer
    
    SwapGem InitialGem, SecondaryGem, Gems()
    Delay 0.5
    HDestroyGems SecondaryGem, Gems(), HValidSwapOne, CurrentScore
    VDestroyGems SecondaryGem, Gems(), VValidSwapOne, CurrentScore
    HDestroyGems InitialGem, Gems(), HValidSwapTwo, CurrentScore
    VDestroyGems InitialGem, Gems(), VValidSwapTwo, CurrentScore
    If VValidSwapOne = False And HValidSwapOne = False And VValidSwapTwo = False And HValidSwapTwo = False Then
        frmBejeweled!tmrIllegalMove.Enabled = True
        SwapGem InitialGem, SecondaryGem, Gems()
    Else
        Delay 0.5
        For k = 0 To 63
            If frmBejeweled!picGem(k).Picture = frmBejeweled!imgGemImage(7).Picture Then
                Gems(k) = 8
            End If
        Next k
        DropGems Gems()
    End If
    
End Sub

'This function destroys any gems in a row of 3 or more.

Sub HDestroyGems(InitialGem As Integer, Gems() As Integer, HValidSwap As Boolean, CurrentScore As Long)

    Dim k As Integer
    Dim j As Integer
    Dim N As Integer
    Dim DestroyGemNum(1 To 5) As Integer

    HValidSwap = True
    N = 3
    DestroyGemNum(1) = -9999

    Select Case InitialGem
        Case 0 To 3, 8 To 11, 16 To 19, 24 To 27, 32 To 35, 40 To 43, 48 To 51, 56 To 59
            If Gems(InitialGem) = Gems(InitialGem + 1) And Gems(InitialGem) = Gems(InitialGem + 2) And Gems(InitialGem) = Gems(InitialGem + 3) And Gems(InitialGem) = Gems(InitialGem + 4) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 1
                DestroyGemNum(3) = InitialGem + 2
                DestroyGemNum(4) = InitialGem + 3
                DestroyGemNum(5) = InitialGem + 4
                N = 5
            End If
    End Select
    
    Select Case InitialGem
        Case 4 To 7, 12 To 15, 20 To 23, 28 To 31, 36 To 39, 44 To 47, 52 To 55, 60 To 63
            If Gems(InitialGem) = Gems(InitialGem - 1) And Gems(InitialGem) = Gems(InitialGem - 2) And Gems(InitialGem) = Gems(InitialGem - 3) And Gems(InitialGem) = Gems(InitialGem - 4) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 1
                DestroyGemNum(3) = InitialGem - 2
                DestroyGemNum(4) = InitialGem - 3
                DestroyGemNum(5) = InitialGem - 4
                N = 5
            End If
    End Select
    
    Select Case InitialGem
        Case 0 To 4, 8 To 12, 16 To 20, 24 To 28, 32 To 36, 40 To 44, 48 To 52, 56 To 60
            If Gems(InitialGem) = Gems(InitialGem + 1) And Gems(InitialGem) = Gems(InitialGem + 2) And Gems(InitialGem) = Gems(InitialGem + 3) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 1
                DestroyGemNum(3) = InitialGem + 2
                DestroyGemNum(4) = InitialGem + 3
                N = 4
            End If
    End Select
    
    Select Case InitialGem
        Case 3 To 7, 11 To 15, 19 To 23, 27 To 31, 35 To 39, 43 To 47, 51 To 55, 59 To 63
            If Gems(InitialGem) = Gems(InitialGem - 1) And Gems(InitialGem) = Gems(InitialGem - 2) And Gems(InitialGem) = Gems(InitialGem - 3) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 1
                DestroyGemNum(3) = InitialGem - 2
                DestroyGemNum(4) = InitialGem - 3
                N = 4
            End If
    End Select
    
    Select Case InitialGem
        Case 0 To 5, 8 To 13, 16 To 21, 24 To 29, 32 To 37, 40 To 45, 48 To 53, 56 To 61
            If Gems(InitialGem) = Gems(InitialGem + 1) And Gems(InitialGem) = Gems(InitialGem + 2) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 1
                DestroyGemNum(3) = InitialGem + 2
            End If
    End Select
    
    Select Case InitialGem
        Case 1 To 6, 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54, 57 To 62
            If Gems(InitialGem) = Gems(InitialGem + 1) And Gems(InitialGem) = Gems(InitialGem - 1) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 1
                DestroyGemNum(3) = InitialGem - 1
            End If
    End Select
    
    Select Case InitialGem
        Case 2 To 7, 10 To 15, 18 To 23, 26 To 31, 34 To 39, 42 To 47, 50 To 55, 58 To 63
            If Gems(InitialGem) = Gems(InitialGem - 1) And Gems(InitialGem) = Gems(InitialGem - 2) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 1
                DestroyGemNum(3) = InitialGem - 2
            End If
    End Select
    
    Select Case InitialGem
        Case 1 To 6, 9 To 14, 17 To 22, 25 To 30, 33 To 38, 41 To 46, 49 To 54, 57 To 62
            If Gems(InitialGem) = Gems(InitialGem - 1) And Gems(InitialGem) = Gems(InitialGem + 1) Then
                Select Case InitialGem
                    Case 1, 9, 17, 25, 33, 41, 49, 57
                        If Gems(InitialGem) = Gems(InitialGem + 2) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem - 1
                            DestroyGemNum(3) = InitialGem + 1
                            DestroyGemNum(4) = InitialGem + 2
                            N = 4
                        End If
                    Case 6, 14, 22, 30, 38, 46, 54, 62
                        If Gems(InitialGem) = Gems(InitialGem - 2) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem - 1
                            DestroyGemNum(3) = InitialGem + 1
                            DestroyGemNum(4) = InitialGem - 2
                            N = 4
                        End If
                    Case Else
                        If Gems(InitialGem) = Gems(InitialGem - 2) And Gems(InitialGem) = Gems(InitialGem + 2) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem - 1
                            DestroyGemNum(3) = InitialGem + 1
                            DestroyGemNum(4) = InitialGem - 2
                            DestroyGemNum(5) = InitialGem + 2
                            N = 5
                        ElseIf Gems(InitialGem) = Gems(InitialGem + 2) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem - 1
                            DestroyGemNum(3) = InitialGem + 1
                            DestroyGemNum(4) = InitialGem + 2
                            N = 4
                        ElseIf Gems(InitialGem) = Gems(InitialGem - 2) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem - 1
                            DestroyGemNum(3) = InitialGem + 1
                            DestroyGemNum(4) = InitialGem - 2
                            N = 4
                        End If
                End Select
            End If
    End Select

    If DestroyGemNum(1) <> -9999 Then
        Select Case N
            Case 3
                CurrentScore = CurrentScore + 100
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
            Case 4
                CurrentScore = CurrentScore + 200
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
            Case 5
                CurrentScore = CurrentScore + 400
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
        End Select
        For j = 1 To N
            frmBejeweled!picGem(DestroyGemNum(j)).Picture = frmBejeweled!imgGemImage(7)
        Next j
    ElseIf DestroyGemNum(1) = -9999 Then
        HValidSwap = False
    End If
    
End Sub

'This general procedure destroys any gems in a column of 3 or more.

Sub VDestroyGems(InitialGem As Integer, Gems() As Integer, VValidSwap As Boolean, CurrentScore As Long)

    Dim k As Integer
    Dim j As Integer
    Dim N As Integer
    Dim DestroyGemNum(1 To 5) As Integer

    N = 3
    VValidSwap = True
    DestroyGemNum(1) = -9999
    
    Select Case InitialGem
        Case 0 To 31
            If Gems(InitialGem) = Gems(InitialGem + 8) And Gems(InitialGem) = Gems(InitialGem + 16) And Gems(InitialGem) = Gems(InitialGem + 24) And Gems(InitialGem) = Gems(InitialGem + 32) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 8
                DestroyGemNum(3) = InitialGem + 16
                DestroyGemNum(4) = InitialGem + 24
                DestroyGemNum(5) = InitialGem + 32
                N = 5
            End If
    End Select
    
    Select Case InitialGem
        Case 32 To 63
            If Gems(InitialGem) = Gems(InitialGem - 8) And Gems(InitialGem) = Gems(InitialGem - 16) And Gems(InitialGem) = Gems(InitialGem - 24) And Gems(InitialGem) = Gems(InitialGem - 32) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 8
                DestroyGemNum(3) = InitialGem - 16
                DestroyGemNum(4) = InitialGem - 24
                DestroyGemNum(5) = InitialGem - 32
                N = 5
            End If
    End Select
    
    Select Case InitialGem
        Case 0 To 39
            If Gems(InitialGem) = Gems(InitialGem + 8) And Gems(InitialGem) = Gems(InitialGem + 16) And Gems(InitialGem) = Gems(InitialGem + 24) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 8
                DestroyGemNum(3) = InitialGem + 16
                DestroyGemNum(4) = InitialGem + 24
                N = 4
            End If
    End Select
    
    Select Case InitialGem
        Case 24 To 63
            If Gems(InitialGem) = Gems(InitialGem - 8) And Gems(InitialGem) = Gems(InitialGem - 16) And Gems(InitialGem) = Gems(InitialGem - 24) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 8
                DestroyGemNum(3) = InitialGem - 16
                DestroyGemNum(4) = InitialGem - 24
                N = 4
            End If
    End Select
    
    Select Case InitialGem
        Case 0 To 47
            If Gems(InitialGem) = Gems(InitialGem + 8) And Gems(InitialGem) = Gems(InitialGem + 16) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 8
                DestroyGemNum(3) = InitialGem + 16
            End If
    End Select
    
    Select Case InitialGem
        Case 16 To 63
            If Gems(InitialGem) = Gems(InitialGem - 8) And Gems(InitialGem) = Gems(InitialGem - 16) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem - 8
                DestroyGemNum(3) = InitialGem - 16
            End If
    End Select
        
    Select Case InitialGem
        Case 8 To 55
            If Gems(InitialGem) = Gems(InitialGem + 8) And Gems(InitialGem) = Gems(InitialGem - 8) Then
                DestroyGemNum(1) = InitialGem
                DestroyGemNum(2) = InitialGem + 8
                DestroyGemNum(3) = InitialGem - 8
            End If
    End Select
    
    Select Case InitialGem
        Case 8 To 55
            If Gems(InitialGem) = Gems(InitialGem + 8) And Gems(InitialGem) = Gems(InitialGem - 8) Then
                Select Case InitialGem
                    Case 8 To 15
                        If Gems(InitialGem) = Gems(InitialGem + 16) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem + 8
                            DestroyGemNum(3) = InitialGem - 8
                            DestroyGemNum(4) = InitialGem + 16
                            N = 4
                        End If
                    Case 48 To 55
                        If Gems(InitialGem) = Gems(InitialGem - 16) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem + 8
                            DestroyGemNum(3) = InitialGem - 8
                            DestroyGemNum(4) = InitialGem - 16
                            N = 4
                        End If
                    Case Else
                        If Gems(InitialGem) = Gems(InitialGem - 16) And Gems(InitialGem) = Gems(InitialGem + 16) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem + 8
                            DestroyGemNum(3) = InitialGem - 8
                            DestroyGemNum(4) = InitialGem - 16
                            DestroyGemNum(5) = InitialGem + 16
                            N = 5
                        ElseIf Gems(InitialGem) = Gems(InitialGem - 16) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem + 8
                            DestroyGemNum(3) = InitialGem - 8
                            DestroyGemNum(4) = InitialGem - 16
                            N = 4
                        ElseIf Gems(InitialGem) = Gems(InitialGem + 16) Then
                            DestroyGemNum(1) = InitialGem
                            DestroyGemNum(2) = InitialGem + 8
                            DestroyGemNum(3) = InitialGem - 8
                            DestroyGemNum(4) = InitialGem + 16
                            N = 4
                        End If
                End Select
            End If
    End Select
    
    If DestroyGemNum(1) <> -9999 Then
        Select Case N
            Case 3
                CurrentScore = CurrentScore + 100
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
            Case 4
                CurrentScore = CurrentScore + 200
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
            Case 5
                CurrentScore = CurrentScore + 400
                frmBejeweled.picScore.Cls
                frmBejeweled.picScore.Print CurrentScore
        End Select
        For j = 1 To N
            frmBejeweled!picGem(DestroyGemNum(j)).Picture = frmBejeweled!imgGemImage(7)
        Next j
    ElseIf DestroyGemNum(1) = -9999 Then
        VValidSwap = False
    End If
    
End Sub

'This general procedure shifts the gems down to fill in the empty spaces.

Sub DropGems(Gem() As Integer)

    Dim k As Integer
    Dim Count As Integer

    For k = 63 To 8 Step -1
        If Gem(k) = 8 Then
            If Gem(k - 8) <> 8 Then
                frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 8).Picture
                frmBejeweled!picGem(k - 8).Picture = frmBejeweled!imgGemImage(7).Picture
                Gem(k) = Gem(k - 8)
                Gem(k - 8) = 8
            ElseIf Gem(k - 8) = 8 Then
                Select Case k
                    Case 32 To 63
                        If Gem(k - 16) <> 8 Then
                            frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 16).Picture
                            frmBejeweled!picGem(k - 16).Picture = frmBejeweled!imgGemImage(7).Picture
                            Gem(k) = Gem(k - 16)
                            Gem(k - 16) = 8
                        ElseIf Gem(k - 16) = 8 Then
                            If Gem(k - 24) <> 8 Then
                                frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 24).Picture
                                frmBejeweled!picGem(k - 24).Picture = frmBejeweled!imgGemImage(7).Picture
                                Gem(k) = Gem(k - 24)
                                Gem(k - 24) = 8
                            ElseIf Gem(k - 24) = 8 Then
                                If Gem(k - 32) <> 8 Then
                                    frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 32).Picture
                                    frmBejeweled!picGem(k - 32).Picture = frmBejeweled!imgGemImage(7).Picture
                                    Gem(k) = Gem(k - 32)
                                    Gem(k - 32) = 8
                                End If
                            End If
                        End If
                    Case 24 To 31
                        If Gem(k - 16) <> 8 Then
                            frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 16).Picture
                            frmBejeweled!picGem(k - 16).Picture = frmBejeweled!imgGemImage(7).Picture
                            Gem(k) = Gem(k - 16)
                            Gem(k - 16) = 8
                        ElseIf Gem(k - 16) = 8 Then
                            If Gem(k - 24) <> 8 Then
                                frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 24).Picture
                                frmBejeweled!picGem(k - 24).Picture = frmBejeweled!imgGemImage(7).Picture
                                Gem(k) = Gem(k - 24)
                                Gem(k - 24) = 8
                            End If
                        End If
                    Case 16 To 23
                        If Gem(k - 16) <> 8 Then
                            frmBejeweled!picGem(k).Picture = frmBejeweled!picGem(k - 16).Picture
                            frmBejeweled!picGem(k - 16).Picture = frmBejeweled!imgGemImage(7).Picture
                            Gem(k) = Gem(k - 16)
                            Gem(k - 16) = 8
                        End If
                End Select
            End If
        End If
    Next k
    
End Sub

'This function generates a single gem.

Function GenerateSingleGem()

    Dim Temp As Integer
    
    Temp = Int(Rnd * 6) + 1
    
    GenerateSingleGem = Temp
    
End Function

'This general procedure determines if there are any gems that need to be destroyed after
'they have been dropped.

Sub DestroyDropGem(Gem() As Integer, CurrentScore As Long)

    Dim k As Integer
    Dim HValidDestroy As Boolean
    Dim VValidDestroy As Boolean
    Dim Counter As Integer
    Dim Num As Integer
    
    Do
        Counter = 0
        For k = 0 To 63
            HDestroyGems k, Gem(), HValidDestroy, CurrentScore
            VDestroyGems k, Gem(), VValidDestroy, CurrentScore
            If HValidDestroy = True Or VValidDestroy = True Then
                Counter = 1
            End If
        Next k
        
        For k = 0 To 63
            If frmBejeweled!picGem(k).Picture = frmBejeweled!imgGemImage(7).Picture Then
                Gem(k) = 8
            End If
        Next k
        
        If Counter = 1 Then
            Delay 0.5
            DropGems Gem()
            For k = 0 To 63
                If Gem(k) = 8 Then
                    Num = GenerateSingleGem
                    Gem(k) = Num
                    frmBejeweled!picGem(k).Picture = frmBejeweled!imgGemImage(Num).Picture
                End If
            Next k
            Delay 0.5
        End If
    Loop While Counter = 1
            
End Sub

'This general procedure enters your high score.

Sub EnterHighScore(CurrentScore As Long, Names() As String, Scores() As Long)

    Dim k As Integer
    Dim j As Integer
    
    k = 0
    
    Do
        k = k + 1
        If CurrentScore > Scores(k) Then
            MsgBox "You got a high score!", vbInformation, "High Score"
            For j = 9 To k Step -1
                Names(j + 1) = Names(j)
                Scores(j + 1) = Scores(j)
            Next j
            Scores(k) = CurrentScore
            Names(k) = InputBox$("Enter your name:", "High Score")
        End If
    Loop While k < 10 And CurrentScore <> Scores(k)
    
    If Scores(k) = CurrentScore Then
        frmHighScore.Cls
        frmHighScore.Print Tab(5); "Name"; Tab(30); "Score"
        frmHighScore.Print
        Open App.Path & "\Bejeweled High Scores.txt" For Output As #1
        
            For k = 1 To 10
                Write #1, Names(k), Scores(k)
                frmHighScore.Print k; Tab(5); Names(k); Tab(30); Scores(k)
            Next k
        
        Close #1
        frmHighScore.Visible = True
    Else
        MsgBox "You did not get a high score.", vbInformation, "No High Score"
        frmBejeweled!cmdQuitGame.Enabled = False
    End If
    
    frmBejeweled!cmdQuitGame.Enabled = False
    
End Sub

'This general procedure determines if you have lost.

Sub IsGameOver(Gem() As Integer, CurrentScore As Long, Names() As String, Scores() As Long)

    Dim k As Integer
    Dim Temp As Boolean
    
    Temp = CheckForMoves(Gem())
    If Temp = False Then
        Delay 2
        For k = 0 To 63
            frmBejeweled!picGem(k).Picture = frmBejeweled!imgGemImage(7).Picture
        Next k
        MsgBox "You have no more moves. You lose.", , "Game Over"
        EnterHighScore CurrentScore, Names(), Scores()
    End If
    
End Sub

'This general procedure causes the program to delay for a certain interval of time.

Sub Delay(Interval As Single)

    Dim T1 As Single
    Dim T2 As Single
    
    T1 = Timer
    
    Do
        T2 = Timer
    Loop While T2 - T1 <> Interval
    
End Sub
