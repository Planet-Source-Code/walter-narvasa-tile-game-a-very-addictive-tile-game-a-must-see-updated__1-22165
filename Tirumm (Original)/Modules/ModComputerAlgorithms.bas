Attribute VB_Name = "mComputerAlgorithms"
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Tirumm: The Tile Rummy Game fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Tirumm: The Tile Rummy Game. Contact me for
' additional help/suggestions
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

Function ValidateTile(xVal As Integer)
    Dim CorrectTrio As Boolean
    Dim P1OmitCond_1 As Boolean, P1OmitCond_2 As Boolean, P1OmitCond_3 As Boolean
    Dim P1OmitCond_4 As Boolean, P1OmitCond_5 As Boolean, P1OmitCond_6 As Boolean
    Dim P1OmitCond_7 As Boolean, P1OmitCond_8 As Boolean, P1OmitCond_9 As Boolean
    Dim P1OmitCond_10 As Boolean, P1OmitCond_11 As Boolean, P1OmitCond_12 As Boolean
    Dim P2OmitCond_1 As Boolean, P2OmitCond_2 As Boolean, P2OmitCond_3 As Boolean
    Dim P2OmitCond_4 As Boolean, P2OmitCond_5 As Boolean, P2OmitCond_6 As Boolean
    Dim P2OmitCond_7 As Boolean, P2OmitCond_8 As Boolean, P2OmitCond_9 As Boolean
    Dim P2OmitCond_10 As Boolean, P2OmitCond_11 As Boolean, P2OmitCond_12 As Boolean
    Dim P3OmitCond_1 As Boolean, P3OmitCond_2 As Boolean, P3OmitCond_3 As Boolean
    Dim P3OmitCond_4 As Boolean, P3OmitCond_5 As Boolean, P3OmitCond_6 As Boolean
    Dim P3OmitCond_7 As Boolean, P3OmitCond_8 As Boolean, P3OmitCond_9 As Boolean
    Dim P3OmitCond_10 As Boolean, P3OmitCond_11 As Boolean, P3OmitCond_12 As Boolean
    If Not GameReset Then
        CorrectTrio = False
        Call Refresh_PlayerTempArrays
        If xVal = 1 Then
            'For i = 0 To 13
            '    MsgBox ("Player One=" & ExArg(1, xPlayer1Temp(i), "-"))
            'Next i
            fMain.cmdDrop.Enabled = False
            fMain.TileStatus.Visible = False
            fMain.GameClockerTime.Visible = False
            fMain.GameClockerCaption.Visible = False
            fMain.GameClocker.Enabled = False
            For i = 0 To 2
                fMain.BufferTile(i).Visible = False
            Next i
            If GameLoopCount = 0 Then
                xP1Val = 0: xP2Val = 0: xP3Val = 0: xP4Val = 0
            Else
                xP1Val = GameLoopCount: xP2Val = GameLoopCount: xP3Val = GameLoopCount: xP4Val = GameLoopCount
            End If
            If Not P1OmitCond_1 Then
                If Val(ExArg(1, xPlayer1Temp(xVT), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 1), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 1), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 2), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT), "-") = ExArg(3, xPlayer1Temp(xVT + 1), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 1), "-") = ExArg(3, xPlayer1Temp(xVT + 2), "-")) Then
                    xP1 = 0: CorrectTrio = True: P1OmitCond_1 = True
                End If
            End If
            If Not P1OmitCond_2 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 2), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 3), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 1), "-") = ExArg(3, xPlayer1Temp(xVT + 2), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 2), "-") = ExArg(3, xPlayer1Temp(xVT + 3), "-")) Then
                    xP1 = 1: CorrectTrio = True: P1OmitCond_2 = True
                End If
            End If
            If Not P1OmitCond_3 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 3), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 4), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 2), "-") = ExArg(3, xPlayer1Temp(xVT + 3), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 3), "-") = ExArg(3, xPlayer1Temp(xVT + 4), "-")) Then
                    xP1 = 2: CorrectTrio = True: P1OmitCond_3 = True
                End If
            End If
            If Not P1OmitCond_4 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 4), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 5), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 3), "-") = ExArg(3, xPlayer1Temp(xVT + 4), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 4), "-") = ExArg(3, xPlayer1Temp(xVT + 5), "-")) Then
                    xP1 = 3: CorrectTrio = True: P1OmitCond_4 = True
                End If
            End If
            If Not P1OmitCond_5 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 5), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 6), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 4), "-") = ExArg(3, xPlayer1Temp(xVT + 5), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 5), "-") = ExArg(3, xPlayer1Temp(xVT + 6), "-")) Then
                    xP1 = 4: CorrectTrio = True: P1OmitCond_5 = True
                End If
            End If
            If Not P1OmitCond_6 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 6), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 7), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 5), "-") = ExArg(3, xPlayer1Temp(xVT + 6), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 6), "-") = ExArg(3, xPlayer1Temp(xVT + 7), "-")) Then
                    xP1 = 5: CorrectTrio = True: P1OmitCond_6 = True
                End If
            End If
            If Not P1OmitCond_7 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 7), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 8), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 6), "-") = ExArg(3, xPlayer1Temp(xVT + 7), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 7), "-") = ExArg(3, xPlayer1Temp(xVT + 8), "-")) Then
                    xP1 = 6: CorrectTrio = True: P1OmitCond_7 = True
                End If
            End If
            If Not P1OmitCond_8 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 8), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 9), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 7), "-") = ExArg(3, xPlayer1Temp(xVT + 8), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 8), "-") = ExArg(3, xPlayer1Temp(xVT + 9), "-")) Then
                    xP1 = 7: CorrectTrio = True: P1OmitCond_8 = True
                End If
            End If
            If Not P1OmitCond_9 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 9), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 10), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 8), "-") = ExArg(3, xPlayer1Temp(xVT + 9), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 9), "-") = ExArg(3, xPlayer1Temp(xVT + 10), "-")) Then
                    xP1 = 8: CorrectTrio = True: P1OmitCond_9 = True
                End If
            End If
            If Not P1OmitCond_10 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 10), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 11), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 9), "-") = ExArg(3, xPlayer1Temp(xVT + 10), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 10), "-") = ExArg(3, xPlayer1Temp(xVT + 11), "-")) Then
                    xP1 = 9: CorrectTrio = True: P1OmitCond_10 = True
                End If
            End If
            If Not P1OmitCond_11 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 12), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 11), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 12), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 10), "-") = ExArg(3, xPlayer1Temp(xVT + 11), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 11), "-") = ExArg(3, xPlayer1Temp(xVT + 12), "-")) Then
                    xP1 = 10: CorrectTrio = True: P1OmitCond_11 = True
                End If
            End If
            If Not P1OmitCond_12 Then
                If Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 12), "-")) And _
                    Val(ExArg(1, xPlayer1Temp(xVT + 12), "-")) = Val(ExArg(1, xPlayer1Temp(xVT + 13), "-")) Or _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 12), "-"))) = _
                    (Val(ExArg(1, xPlayer1Temp(xVT + 12), "-")) - Val(ExArg(1, xPlayer1Temp(xVT + 13), "-"))) And _
                    (ExArg(3, xPlayer1Temp(xVT + 11), "-") = ExArg(3, xPlayer1Temp(xVT + 12), "-") And _
                    ExArg(3, xPlayer1Temp(xVT + 12), "-") = ExArg(3, xPlayer1Temp(xVT + 13), "-")) Then
                    xP1 = 11: CorrectTrio = True: P1OmitCond_12 = True
                End If
            End If
            If CorrectTrio = True Then
                'MsgBox ("Player One xDpVal=>" & xDpVal)
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Visible = False
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Visible = False
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Visible = False
                fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Visible = False
                fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Visible = False
                fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Visible = False
                xPlayer1Out(xP1Val) = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).ToolTipText)
                xPlayer1Out(xP1Val + 1) = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).ToolTipText)
                xPlayer1Out(xP1Val + 2) = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).ToolTipText)
                fMain.DropTile(xDpVal).Caption = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Caption)
                fMain.DropTile(xDpVal).BackColor = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).BackColor
                fMain.DropTile(xDpVal).ToolTipText = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 1).Caption = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Caption)
                fMain.DropTile(xDpVal + 1).BackColor = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).BackColor
                fMain.DropTile(xDpVal + 1).ToolTipText = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 2).Caption = Val(fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Caption)
                fMain.DropTile(xDpVal + 2).BackColor = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).BackColor
                fMain.DropTile(xDpVal + 2).ToolTipText = fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).ToolTipText
                xDpVal = xDpVal + 3
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Caption = ""
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Caption = ""
                fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Caption = ""
                FillInTiles = True
                Call Random_Pick(1)
                If Not GameReset Then
                    'MsgBox "Congratulations! Player 1 has completed a combination of Trio." & vbCrLf & _
                    '        "Player 2 will make the next move.", vbOKOnly + vbInformation, "Player 1's Alert"
                    fMain.MoveStatus.AddItem "Congratulations! Player 1 has completed a tile combination."
                    fMain.MoveStatus.AddItem "Player 2 will make the next move."
                    fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                    If SoundOn = True Then
                        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
                    End If
                    fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Visible = True
                    fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Visible = True
                    fMain.AIPlayer1Tile(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Visible = True
                    fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1), "-"))).Visible = True
                    fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1 + 1), "-"))).Visible = True
                    fMain.AIPlayer1Cover(Val(ExArg(2, xPlayer1Temp(xP1 + 2), "-"))).Visible = True
                    GameLoopCount = GameLoopCount + 1
                Else
                    Call fMain.ClearAllTiles
                End If
            Else
                'MsgBox "Player 1 has no move.", vbOKOnly + vbInformation, "Player 1's Alert"
                fMain.MoveStatus.AddItem "Player 1 has no move."
                fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                If SoundOn = True Then
                    Call sndPlaySound(App.Path & "\Sounds\Alert.wav", 1)
                End If
                Call Force_PickTiles(1)
            End If
            CurrentTurn = "P2"
            fMain.TurnStatus.Caption = "Player 2's Turn"
            Call ValidateTile(2)
        ElseIf xVal = 2 Then
            'For i = 0 To 13
            '    MsgBox ("Player Two=" & ExArg(1, xPlayer2Temp(i), "-"))
            'Next i
            fMain.cmdDrop.Enabled = False
            fMain.TileStatus.Visible = False
            fMain.GameClockerTime.Visible = False
            fMain.GameClockerCaption.Visible = False
            fMain.GameClocker.Enabled = False
            For i = 0 To 2
                fMain.BufferTile(i).Visible = False
            Next i
            If GameLoopCount = 0 Then
                xP1Val = 0: xP2Val = 0: xP3Val = 0: xP4Val = 0
            Else
                xP1Val = GameLoopCount: xP2Val = GameLoopCount: xP3Val = GameLoopCount: xP4Val = GameLoopCount
            End If
            If Not P2OmitCond_1 Then
                If Val(ExArg(1, xPlayer2Temp(xVT), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 1), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 1), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 2), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT), "-") = ExArg(3, xPlayer2Temp(xVT + 1), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 1), "-") = ExArg(3, xPlayer2Temp(xVT + 2), "-")) Then
                    xP2 = 0: CorrectTrio = True: P2OmitCond_1 = True
                End If
            End If
            If Not P2OmitCond_2 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 2), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 3), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 1), "-") = ExArg(3, xPlayer2Temp(xVT + 2), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 2), "-") = ExArg(3, xPlayer2Temp(xVT + 3), "-")) Then
                    xP2 = 1: CorrectTrio = True: P2OmitCond_2 = True
                End If
            End If
            If Not P2OmitCond_3 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 3), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 4), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 2), "-") = ExArg(3, xPlayer2Temp(xVT + 3), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 3), "-") = ExArg(3, xPlayer2Temp(xVT + 4), "-")) Then
                    xP2 = 2: CorrectTrio = True: P2OmitCond_3 = True
                End If
            End If
            If Not P2OmitCond_4 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 4), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 5), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 3), "-") = ExArg(3, xPlayer2Temp(xVT + 4), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 4), "-") = ExArg(3, xPlayer2Temp(xVT + 5), "-")) Then
                    xP2 = 3: CorrectTrio = True: P2OmitCond_4 = True
                End If
            End If
            If Not P2OmitCond_5 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 5), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 6), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 4), "-") = ExArg(3, xPlayer2Temp(xVT + 5), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 5), "-") = ExArg(3, xPlayer2Temp(xVT + 6), "-")) Then
                    xP2 = 4: CorrectTrio = True: P2OmitCond_5 = True
                End If
            End If
            If Not P2OmitCond_6 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 6), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 7), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 5), "-") = ExArg(3, xPlayer2Temp(xVT + 6), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 6), "-") = ExArg(3, xPlayer2Temp(xVT + 7), "-")) Then
                    xP2 = 5: CorrectTrio = True: P2OmitCond_6 = True
                End If
            End If
            If Not P2OmitCond_7 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 7), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 8), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 6), "-") = ExArg(3, xPlayer2Temp(xVT + 7), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 7), "-") = ExArg(3, xPlayer2Temp(xVT + 8), "-")) Then
                    xP2 = 6: CorrectTrio = True: P2OmitCond_7 = True
                End If
            End If
            If Not P2OmitCond_8 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 8), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 9), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 7), "-") = ExArg(3, xPlayer2Temp(xVT + 8), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 8), "-") = ExArg(3, xPlayer2Temp(xVT + 9), "-")) Then
                    xP2 = 7: CorrectTrio = True: P2OmitCond_8 = True
                End If
            End If
            If Not P2OmitCond_9 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 9), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 10), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 8), "-") = ExArg(3, xPlayer2Temp(xVT + 9), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 9), "-") = ExArg(3, xPlayer2Temp(xVT + 10), "-")) Then
                    xP2 = 8: CorrectTrio = True: P2OmitCond_9 = True
                End If
            End If
            If Not P2OmitCond_10 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 10), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 11), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 9), "-") = ExArg(3, xPlayer2Temp(xVT + 10), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 10), "-") = ExArg(3, xPlayer2Temp(xVT + 11), "-")) Then
                    xP2 = 9: CorrectTrio = True: P2OmitCond_10 = True
                End If
            End If
            If Not P2OmitCond_11 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 12), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 11), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 12), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 10), "-") = ExArg(3, xPlayer2Temp(xVT + 11), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 11), "-") = ExArg(3, xPlayer2Temp(xVT + 12), "-")) Then
                    xP2 = 10: CorrectTrio = True: P2OmitCond_11 = True
                End If
            End If
            If Not P2OmitCond_12 Then
                If Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 12), "-")) And _
                    Val(ExArg(1, xPlayer2Temp(xVT + 12), "-")) = Val(ExArg(1, xPlayer2Temp(xVT + 13), "-")) Or _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 12), "-"))) = _
                    (Val(ExArg(1, xPlayer2Temp(xVT + 12), "-")) - Val(ExArg(1, xPlayer2Temp(xVT + 13), "-"))) And _
                    (ExArg(3, xPlayer2Temp(xVT + 11), "-") = ExArg(3, xPlayer2Temp(xVT + 12), "-") And _
                    ExArg(3, xPlayer2Temp(xVT + 12), "-") = ExArg(3, xPlayer2Temp(xVT + 13), "-")) Then
                    xP2 = 11: CorrectTrio = True: P2OmitCond_12 = True
                End If
            End If
            If CorrectTrio = True Then
                'MsgBox ("Player Two xDpVal=>" & xDpVal)
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Visible = False
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Visible = False
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Visible = False
                fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Visible = False
                fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Visible = False
                fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Visible = False
                xPlayer2Out(xP2Val) = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).ToolTipText)
                xPlayer2Out(xP2Val + 1) = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).ToolTipText)
                xPlayer2Out(xP2Val + 2) = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).ToolTipText)
                fMain.DropTile(xDpVal).Caption = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Caption)
                fMain.DropTile(xDpVal).BackColor = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).BackColor
                fMain.DropTile(xDpVal).ToolTipText = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 1).Caption = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Caption)
                fMain.DropTile(xDpVal + 1).BackColor = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).BackColor
                fMain.DropTile(xDpVal + 1).ToolTipText = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 2).Caption = Val(fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Caption)
                fMain.DropTile(xDpVal + 2).BackColor = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).BackColor
                fMain.DropTile(xDpVal + 2).ToolTipText = fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).ToolTipText
                xDpVal = xDpVal + 3
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Caption = ""
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Caption = ""
                fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Caption = ""
                FillInTiles = True
                Call Random_Pick(2)
                If Not GameReset Then
                    'MsgBox "Congratulations! Player 2 has completed a combination of Trio." & vbCrLf & _
                    '        "Player 3 will make the next move.", vbOKOnly + vbInformation, "Player 2's Alert"
                    fMain.MoveStatus.AddItem "Congratulations! Player 2 has completed a tile combination."
                    fMain.MoveStatus.AddItem "Player 3 will make the next move."
                    fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                    If SoundOn = True Then
                        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
                    End If
                    fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Visible = True
                    fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Visible = True
                    fMain.AIPlayer2Tile(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Visible = True
                    fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2), "-"))).Visible = True
                    fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2 + 1), "-"))).Visible = True
                    fMain.AIPlayer2Cover(Val(ExArg(2, xPlayer2Temp(xP2 + 2), "-"))).Visible = True
                    GameLoopCount = GameLoopCount + 1
                Else
                    Call fMain.ClearAllTiles
                End If
            Else
                'MsgBox "Player 2 has no move.", vbOKOnly + vbInformation, "Player 2's Alert"
                fMain.MoveStatus.AddItem "Player 2 has no move."
                fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                If SoundOn = True Then
                    Call sndPlaySound(App.Path & "\Sounds\Alert.wav", 1)
                End If
                Call Force_PickTiles(2)
            End If
            CurrentTurn = "P3"
            fMain.TurnStatus.Caption = "Player 3's Turn"
            Call ValidateTile(3)
        ElseIf xVal = 3 Then
            'For i = 0 To 13
            '    MsgBox ("Player Three=" & ExArg(1, xPlayer3Temp(i), "-"))
            'Next i
            fMain.cmdDrop.Enabled = False
            fMain.TileStatus.Visible = False
            fMain.GameClockerTime.Visible = False
            fMain.GameClockerCaption.Visible = False
            fMain.GameClocker.Enabled = False
            For i = 0 To 2
                fMain.BufferTile(i).Visible = False
            Next i
            If GameLoopCount = 0 Then
                xP1Val = 0: xP2Val = 0: xP3Val = 0: xP4Val = 0
            Else
                xP1Val = GameLoopCount: xP2Val = GameLoopCount: xP3Val = GameLoopCount: xP4Val = GameLoopCount
            End If
            If Not P3OmitCond_1 Then
                If Val(ExArg(1, xPlayer3Temp(xVT), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 1), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 1), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 2), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT), "-") = ExArg(3, xPlayer3Temp(xVT + 1), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 1), "-") = ExArg(3, xPlayer3Temp(xVT + 2), "-")) Then
                    xP3 = 0: CorrectTrio = True: P3OmitCond_1 = True
                End If
            End If
            If Not P3OmitCond_2 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 1), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 1), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 2), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 3), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 1), "-") = ExArg(3, xPlayer3Temp(xVT + 2), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 2), "-") = ExArg(3, xPlayer3Temp(xVT + 3), "-")) Then
                    xP3 = 1: CorrectTrio = True: P3OmitCond_2 = True
                End If
            End If
            If Not P3OmitCond_3 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 2), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 3), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 4), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 2), "-") = ExArg(3, xPlayer3Temp(xVT + 3), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 3), "-") = ExArg(3, xPlayer3Temp(xVT + 4), "-")) Then
                    xP3 = 2: CorrectTrio = True: P3OmitCond_3 = True
                End If
            End If
            If Not P3OmitCond_4 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 3), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 4), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 5), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 3), "-") = ExArg(3, xPlayer3Temp(xVT + 4), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 4), "-") = ExArg(3, xPlayer3Temp(xVT + 5), "-")) Then
                    xP3 = 3: CorrectTrio = True: P3OmitCond_4 = True
                End If
            End If
            If Not P3OmitCond_5 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 4), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 5), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 6), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 4), "-") = ExArg(3, xPlayer3Temp(xVT + 5), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 5), "-") = ExArg(3, xPlayer3Temp(xVT + 6), "-")) Then
                    xP3 = 4: CorrectTrio = True: P3OmitCond_5 = True
                End If
            End If
            If Not P3OmitCond_6 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 5), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 6), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 7), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 5), "-") = ExArg(3, xPlayer3Temp(xVT + 6), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 6), "-") = ExArg(3, xPlayer3Temp(xVT + 7), "-")) Then
                    xP3 = 5: CorrectTrio = True: P3OmitCond_6 = True
                End If
            End If
            If Not P3OmitCond_7 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 6), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 7), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 8), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 6), "-") = ExArg(3, xPlayer3Temp(xVT + 7), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 7), "-") = ExArg(3, xPlayer3Temp(xVT + 8), "-")) Then
                    xP3 = 6: CorrectTrio = True: P3OmitCond_7 = True
                End If
            End If
            If Not P3OmitCond_8 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 7), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 8), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 9), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 7), "-") = ExArg(3, xPlayer3Temp(xVT + 8), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 8), "-") = ExArg(3, xPlayer3Temp(xVT + 9), "-")) Then
                    xP3 = 7: CorrectTrio = True: P3OmitCond_8 = True
                End If
            End If
            If Not P3OmitCond_9 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 8), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 9), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 10), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 8), "-") = ExArg(3, xPlayer3Temp(xVT + 9), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 9), "-") = ExArg(3, xPlayer3Temp(xVT + 10), "-")) Then
                    xP3 = 8: CorrectTrio = True: P3OmitCond_9 = True
                End If
            End If
            If Not P3OmitCond_10 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 9), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 10), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 11), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 9), "-") = ExArg(3, xPlayer3Temp(xVT + 10), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 10), "-") = ExArg(3, xPlayer3Temp(xVT + 11), "-")) Then
                    xP3 = 9: CorrectTrio = True: P3OmitCond_10 = True
                End If
            End If
            If Not P3OmitCond_11 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 12), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 10), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 11), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 12), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 10), "-") = ExArg(3, xPlayer3Temp(xVT + 11), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 11), "-") = ExArg(3, xPlayer3Temp(xVT + 12), "-")) Then
                    xP3 = 10: CorrectTrio = True: P3OmitCond_11 = True
                End If
            End If
            If Not P3OmitCond_12 Then
                If Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 12), "-")) And _
                    Val(ExArg(1, xPlayer3Temp(xVT + 12), "-")) = Val(ExArg(1, xPlayer3Temp(xVT + 13), "-")) Or _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 11), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 12), "-"))) = _
                    (Val(ExArg(1, xPlayer3Temp(xVT + 12), "-")) - Val(ExArg(1, xPlayer3Temp(xVT + 13), "-"))) And _
                    (ExArg(3, xPlayer3Temp(xVT + 11), "-") = ExArg(3, xPlayer3Temp(xVT + 12), "-") And _
                    ExArg(3, xPlayer3Temp(xVT + 12), "-") = ExArg(3, xPlayer3Temp(xVT + 13), "-")) Then
                    xP3 = 11: CorrectTrio = True: P3OmitCond_12 = True
                End If
            End If
            If CorrectTrio = True Then
                'MsgBox ("Player Three xDpVal=>" & xDpVal)
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Visible = False
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Visible = False
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Visible = False
                fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Visible = False
                fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Visible = False
                fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Visible = False
                xPlayer3Out(xP3Val) = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).ToolTipText)
                xPlayer3Out(xP3Val + 1) = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).ToolTipText)
                xPlayer3Out(xP3Val + 2) = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).ToolTipText)
                fMain.DropTile(xDpVal).Caption = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Caption)
                fMain.DropTile(xDpVal).BackColor = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).BackColor
                fMain.DropTile(xDpVal).ToolTipText = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 1).Caption = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Caption)
                fMain.DropTile(xDpVal + 1).BackColor = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).BackColor
                fMain.DropTile(xDpVal + 1).ToolTipText = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).ToolTipText
                fMain.DropTile(xDpVal + 2).Caption = Val(fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Caption)
                fMain.DropTile(xDpVal + 2).BackColor = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).BackColor
                fMain.DropTile(xDpVal + 2).ToolTipText = fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).ToolTipText
                xDpVal = xDpVal + 3
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Caption = ""
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Caption = ""
                fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Caption = ""
                FillInTiles = True
                Call Random_Pick(3)
                If Not GameReset Then
                    'MsgBox "Congratulations! Player 3 has completed a combination of Trio." & vbCrLf & _
                    '        "Player 4 will make the next move.", vbOKOnly + vbInformation, "Player 3's Alert"
                    fMain.MoveStatus.AddItem "Congratulations! Player 3 has completed a tile combination."
                    fMain.MoveStatus.AddItem "Player 4 will make the next move."
                    fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                    If SoundOn = True Then
                        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
                    End If
                    fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Visible = True
                    fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Visible = True
                    fMain.AIPlayer3Tile(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Visible = True
                    fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3), "-"))).Visible = True
                    fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3 + 1), "-"))).Visible = True
                    fMain.AIPlayer3Cover(Val(ExArg(2, xPlayer3Temp(xP3 + 2), "-"))).Visible = True
                    GameLoopCount = GameLoopCount + 1
                Else
                    Call fMain.ClearAllTiles
                End If
            Else
                'MsgBox "Player 3 has no move.", vbOKOnly + vbInformation, "Player 3's Alert"
                fMain.MoveStatus.AddItem "Player 3 has no move."
                fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                If SoundOn = True Then
                    Call sndPlaySound(App.Path & "\Sounds\Alert.wav", 1)
                End If
                Call Force_PickTiles(3)
            End If
            fMain.cmdDrop.Enabled = True
            fMain.TileStatus.Visible = True
            fMain.GameClockerTime.Visible = True
            fMain.GameClockerCaption.Visible = True
            fMain.GameClocker.Enabled = True
            For i = 0 To 2
                fMain.BufferTile(i).Visible = True
            Next i
            CurrentTurn = "P4"
            fMain.TurnStatus.Caption = "Player 4's Turn"
            Call ValidateTile(4)
        End If
    End If
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Effects.wav", 1)
    End If
End Function

' RANDOM PICK TILES & DISTRIBUTE TO ALL PLAYERS->1/2/3/4 LINE UP TILES
Function Random_Pick(xVal As Integer)
    Dim P1Ctr As Integer, P2Ctr As Integer, P3Ctr As Integer, P4Ctr As Integer
    P1Ctr = 0: P2Ctr = 0: P3Ctr = 0: P4Ctr = 0
    'If Not GameReset Then
        ' PLAYER 1 RANDOM PICK TILES FROM CENTER ISLE
        If xVal = 1 Then
            Do While P1Ctr <> 14
                For i = 0 To 13
                    If fMain.AIPlayer1Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer1Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer1Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer1Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P1Ctr = P1Ctr + 1
                            'MsgBox (P1Ctr)
                        End If
                    End If
                Next i
                If FillInTiles = True Then
                    If P1Ctr = 3 Then FillInTiles = False: Exit Function
                End If
                If ValidateWinner = True Then
                    If P1Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 1's Alert RANDOM PICK"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ' PLAYER 2 RANDOM PICK TILES FROM CENTER ISLE
        ElseIf xVal = 2 Then
            Do While P2Ctr <> 14
                For i = 0 To 13
                    If fMain.AIPlayer2Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer2Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer2Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer2Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P2Ctr = P2Ctr + 1
                            'MsgBox (P2Ctr)
                        End If
                    End If
                Next i
                If FillInTiles = True Then
                    If P2Ctr = 3 Then FillInTiles = False: Exit Function
                End If
                If ValidateWinner = True Then
                    If P2Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 2's Alert RANDOM PICK"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ' PLAYER 3 RANDOM PICK TILES FROM CENTER ISLE
        ElseIf xVal = 3 Then
            Do While P3Ctr <> 14
                For i = 0 To 13
                    If fMain.AIPlayer3Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer3Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer3Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer3Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P3Ctr = P3Ctr + 1
                            'MsgBox (P3Ctr)
                        End If
                    End If
                Next i
                If FillInTiles = True Then
                    If P3Ctr = 3 Then FillInTiles = False: Exit Function
                End If
                If ValidateWinner = True Then
                    If P3Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 3's Alert RANDOM PICK"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ' PLAYER 4 RANDOM PICK TILES FROM CENTER ISLE
        ElseIf xVal = 4 Then
            Do While P4Ctr <> 14
                For i = 0 To 13
                    If fMain.AIPlayer4Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer4Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer4Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer4Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P4Ctr = P4Ctr + 1
                            'MsgBox (P4Ctr)
                        End If
                    End If
                Next i
                If FillInTiles = True Then
                    If P4Ctr = 3 Then FillInTiles = False: Exit Function
                End If
                If ValidateWinner = True Then
                    If P4Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 4's Alert RANDOM PICK"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        End If
    'End If
End Function

'REFRESH ALL PLAYERS TEMPORARY ARRAYS
Function Refresh_PlayerTempArrays()
    For a = 0 To 13
        xPlayer1Temp(a) = fMain.AIPlayer1Tile(a).Caption & "-" & fMain.AIPlayer1Tile(a).Index & _
                        "-" & fMain.AIPlayer1Tile(a).BackColor
    Next a
    Call SortStringArray(xPlayer1Temp(), False)
    For b = 0 To 13
        xPlayer2Temp(b) = fMain.AIPlayer2Tile(b).Caption & "-" & fMain.AIPlayer2Tile(b).Index & _
                        "-" & fMain.AIPlayer2Tile(b).BackColor
    Next b
    Call SortStringArray(xPlayer2Temp(), False)
    For c = 0 To 13
        xPlayer3Temp(c) = fMain.AIPlayer3Tile(c).Caption & "-" & fMain.AIPlayer3Tile(c).Index & _
                        "-" & fMain.AIPlayer3Tile(c).BackColor
    Next c
    Call SortStringArray(xPlayer3Temp(), False)
End Function

' DETERMINES WHO IS THE CURRENT WINNER IN A ROUND
Function ValidateWinner() As Boolean
    Dim yCount As Integer, yTile As Integer
    yCount = 0: yTile = 0
    For y = 0 To 105
        If fMain.Tiles(y).Visible = False Then
            If yTile = 103 Then
                ValidateWinner = True
                Exit Function
            End If
            yTile = yTile + 1
        End If
        yCount = yCount + 1
    Next y
End Function

' METHOD IN WHICH THE PLAYER IS FORCE TO PICK FROM CENTER ISLE TILES
Function Force_PickTiles(xVal As Integer)
    If xVal = 1 Then
        For x = 0 To 27
            If fMain.AIPlayer1Tile(x).Caption = "" Then
                If xP1Ctr = 3 Then
                    Call Random_Pick2(1, 14, 16)
                    fMain.AIPlayer1Tile(14).Visible = True: fMain.AIPlayer1Tile(15).Visible = True: fMain.AIPlayer1Tile(16).Visible = True
                    fMain.AIPlayer1Cover(14).Visible = True: fMain.AIPlayer1Cover(15).Visible = True: fMain.AIPlayer1Cover(16).Visible = True
                    xP1Ctr = xP1Ctr + 1
                    Exit Function
                ElseIf xP1Ctr = 4 Then
                    Call Random_Pick2(1, 17, 19)
                    fMain.AIPlayer1Tile(17).Visible = True: fMain.AIPlayer1Tile(18).Visible = True: fMain.AIPlayer1Tile(19).Visible = True:
                    fMain.AIPlayer1Cover(17).Visible = True: fMain.AIPlayer1Cover(18).Visible = True: fMain.AIPlayer1Cover(19).Visible = True
                    xP1Ctr = xP1Ctr + 1
                    Exit Function
                ElseIf xP1Ctr = 5 Then
                    Call Random_Pick2(1, 20, 22)
                    fMain.AIPlayer1Tile(20).Visible = True: fMain.AIPlayer1Tile(21).Visible = True: fMain.AIPlayer1Tile(22).Visible = True
                    fMain.AIPlayer1Cover(20).Visible = True: fMain.AIPlayer1Cover(21).Visible = True: fMain.AIPlayer1Cover(22).Visible = True
                    xP1Ctr = xP1Ctr + 1
                    Exit Function
                ElseIf xP1Ctr = 6 Then
                    Call Random_Pick2(1, 23, 25)
                    fMain.AIPlayer1Tile(23).Visible = True: fMain.AIPlayer1Tile(24).Visible = True: fMain.AIPlayer1Tile(25).Visible = True
                    fMain.AIPlayer1Cover(23).Visible = True: fMain.AIPlayer1Cover(24).Visible = True: fMain.AIPlayer1Cover(25).Visible = True
                    xP1Ctr = xP1Ctr + 1
                    Exit Function
                ElseIf xP1Ctr = 7 Then
                    Call Random_Pick2(1, 26, 27)
                    fMain.AIPlayer1Tile(26).Visible = True: fMain.AIPlayer1Tile(27).Visible = True
                    fMain.AIPlayer1Cover(26).Visible = True: fMain.AIPlayer1Cover(27).Visible = True
                    xP1Ctr = xP1Ctr + 1
                    Exit Function
                End If
                xP1Ctr = xP1Ctr + 1
            End If
        Next x
    ElseIf xVal = 2 Then
        For x = 0 To 27
            If fMain.AIPlayer2Tile(x).Caption = "" Then
                If xP2Ctr = 3 Then
                    Call Random_Pick2(2, 14, 16)
                    fMain.AIPlayer2Tile(14).Visible = True: fMain.AIPlayer2Tile(15).Visible = True: fMain.AIPlayer2Tile(16).Visible = True
                    fMain.AIPlayer2Cover(14).Visible = True: fMain.AIPlayer2Cover(15).Visible = True: fMain.AIPlayer2Cover(16).Visible = True
                    xP2Ctr = xP2Ctr + 1
                    Exit Function
                ElseIf xP2Ctr = 4 Then
                    Call Random_Pick2(2, 17, 19)
                    fMain.AIPlayer2Tile(17).Visible = True: fMain.AIPlayer2Tile(18).Visible = True: fMain.AIPlayer2Tile(19).Visible = True:
                    fMain.AIPlayer2Cover(17).Visible = True: fMain.AIPlayer2Cover(18).Visible = True: fMain.AIPlayer2Cover(19).Visible = True
                    xP2Ctr = xP2Ctr + 1
                    Exit Function
                ElseIf xP2Ctr = 5 Then
                    Call Random_Pick2(2, 20, 22)
                    fMain.AIPlayer2Tile(20).Visible = True: fMain.AIPlayer2Tile(21).Visible = True: fMain.AIPlayer2Tile(22).Visible = True
                    fMain.AIPlayer2Cover(20).Visible = True: fMain.AIPlayer2Cover(21).Visible = True: fMain.AIPlayer2Cover(22).Visible = True
                    xP2Ctr = xP2Ctr + 1
                    Exit Function
                ElseIf xP2Ctr = 6 Then
                    Call Random_Pick2(2, 23, 25)
                    fMain.AIPlayer2Tile(23).Visible = True: fMain.AIPlayer2Tile(24).Visible = True: fMain.AIPlayer2Tile(25).Visible = True
                    fMain.AIPlayer2Cover(23).Visible = True: fMain.AIPlayer2Cover(24).Visible = True: fMain.AIPlayer2Cover(25).Visible = True
                    xP2Ctr = xP2Ctr + 1
                    Exit Function
                ElseIf xP2Ctr = 7 Then
                    Call Random_Pick2(2, 26, 27)
                    fMain.AIPlayer2Tile(26).Visible = True: fMain.AIPlayer2Tile(27).Visible = True
                    fMain.AIPlayer2Cover(26).Visible = True: fMain.AIPlayer2Cover(27).Visible = True
                    xP2Ctr = xP2Ctr + 1
                    Exit Function
                End If
                xP2Ctr = xP2Ctr + 1
            End If
        Next x
    ElseIf xVal = 3 Then
        For x = 0 To 27
            If fMain.AIPlayer3Tile(x).Caption = "" Then
                If xP3Ctr = 3 Then
                    Call Random_Pick2(3, 14, 16)
                    fMain.AIPlayer3Tile(14).Visible = True: fMain.AIPlayer3Tile(15).Visible = True: fMain.AIPlayer3Tile(16).Visible = True
                    fMain.AIPlayer3Cover(14).Visible = True: fMain.AIPlayer3Cover(15).Visible = True: fMain.AIPlayer3Cover(16).Visible = True
                    xP3Ctr = xP3Ctr + 1
                    Exit Function
                ElseIf xP3Ctr = 4 Then
                    Call Random_Pick2(3, 17, 19)
                    fMain.AIPlayer3Tile(17).Visible = True: fMain.AIPlayer3Tile(18).Visible = True: fMain.AIPlayer3Tile(19).Visible = True:
                    fMain.AIPlayer3Cover(17).Visible = True: fMain.AIPlayer3Cover(18).Visible = True: fMain.AIPlayer3Cover(19).Visible = True
                    xP3Ctr = xP3Ctr + 1
                    Exit Function
                ElseIf xP3Ctr = 5 Then
                    Call Random_Pick2(3, 20, 22)
                    fMain.AIPlayer3Tile(20).Visible = True: fMain.AIPlayer3Tile(21).Visible = True: fMain.AIPlayer3Tile(22).Visible = True
                    fMain.AIPlayer3Cover(20).Visible = True: fMain.AIPlayer3Cover(21).Visible = True: fMain.AIPlayer3Cover(22).Visible = True
                    xP3Ctr = xP3Ctr + 1
                    Exit Function
                ElseIf xP3Ctr = 6 Then
                    Call Random_Pick2(3, 23, 25)
                    fMain.AIPlayer3Tile(23).Visible = True: fMain.AIPlayer3Tile(24).Visible = True: fMain.AIPlayer3Tile(25).Visible = True
                    fMain.AIPlayer3Cover(23).Visible = True: fMain.AIPlayer3Cover(24).Visible = True: fMain.AIPlayer3Cover(25).Visible = True
                    xP3Ctr = xP3Ctr + 1
                    Exit Function
                ElseIf xP3Ctr = 7 Then
                    Call Random_Pick2(3, 26, 27)
                    fMain.AIPlayer3Tile(26).Visible = True: fMain.AIPlayer3Tile(27).Visible = True
                    fMain.AIPlayer3Cover(26).Visible = True: fMain.AIPlayer3Cover(27).Visible = True
                    xP3Ctr = xP3Ctr + 1
                    Exit Function
                End If
                xP3Ctr = xP3Ctr + 1
            End If
        Next x
    ElseIf xVal = 4 Then
        For x = 0 To 27
            If fMain.AIPlayer4Tile(x).Caption = "" Then
                If xP4Ctr = 3 Then
                    Call Random_Pick2(4, 14, 16)
                    fMain.AIPlayer4Tile(14).Visible = True: fMain.AIPlayer4Tile(15).Visible = True: fMain.AIPlayer4Tile(16).Visible = True
                    xP4Ctr = xP4Ctr + 1
                    Exit Function
                ElseIf xP4Ctr = 4 Then
                    Call Random_Pick2(4, 17, 19)
                    fMain.AIPlayer4Tile(17).Visible = True: fMain.AIPlayer4Tile(18).Visible = True: fMain.AIPlayer4Tile(19).Visible = True:
                    xP4Ctr = xP4Ctr + 1
                    Exit Function
                ElseIf xP4Ctr = 5 Then
                    Call Random_Pick2(4, 20, 22)
                    fMain.AIPlayer4Tile(20).Visible = True: fMain.AIPlayer4Tile(21).Visible = True: fMain.AIPlayer4Tile(22).Visible = True
                    xP4Ctr = xP4Ctr + 1
                    Exit Function
                ElseIf xP4Ctr = 6 Then
                    Call Random_Pick2(4, 23, 25)
                    fMain.AIPlayer4Tile(23).Visible = True: fMain.AIPlayer4Tile(24).Visible = True: fMain.AIPlayer4Tile(25).Visible = True
                    xP4Ctr = xP4Ctr + 1
                    Exit Function
                ElseIf xP4Ctr = 7 Then
                    Call Random_Pick2(4, 26, 27)
                    fMain.AIPlayer4Tile(26).Visible = True: fMain.AIPlayer4Tile(27).Visible = True
                    xP4Ctr = xP4Ctr + 1
                    Exit Function
                End If
                xP4Ctr = xP4Ctr + 1
            End If
        Next x
    End If
End Function

' RANDOM PICK VERSION 2 TILES & DISTRIBUTE TO ALL PLAYERS->1/2/3/4 2ND LINE UP TILES
Function Random_Pick2(xVal As Integer, xPoint As Integer, yPoint As Integer)
    Dim P1Ctr As Integer, P2Ctr As Integer, P3Ctr As Integer, P4Ctr As Integer
    P1Ctr = 0: P2Ctr = 0: P3Ctr = 0: P4Ctr = 0
    'If Not GameReset Then
        If xVal = 1 Then
            Do While P1Ctr <> 3
                For i = xPoint To yPoint
                    If fMain.AIPlayer1Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer1Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer1Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer1Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P1Ctr = P1Ctr + 1
                        End If
                    End If
                Next i
                If ValidateWinner = True Then
                    If P1Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 1's Alert RANDOM PICK 2"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ElseIf xVal = 2 Then
            Do While P2Ctr <> 3
                For i = xPoint To yPoint
                    If fMain.AIPlayer2Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer2Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer2Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer2Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P2Ctr = P2Ctr + 1
                        End If
                    End If
                Next i
                If ValidateWinner = True Then
                    If P2Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 2's Alert  RANDOM PICK 2"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ElseIf xVal = 3 Then
            Do While P3Ctr <> 3
                For i = xPoint To yPoint
                    If fMain.AIPlayer3Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer3Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer3Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer3Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P3Ctr = P3Ctr + 1
                        End If
                    End If
                Next i
                If ValidateWinner = True Then
                    If P3Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 3's Alert  RANDOM PICK 2"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        ElseIf xVal = 4 Then
            Do While P4Ctr <> 3
                For i = xPoint To yPoint
                    If fMain.AIPlayer4Tile(i).Caption = "" Then
                        Randomize
                        x = Int(Rnd(1) * ((106 - 1) + 1))
                        If fMain.Tiles(x).Visible = True Then
                            fMain.AIPlayer4Tile(i).Caption = fMain.Tiles(x).Caption
                            fMain.AIPlayer4Tile(i).BackColor = fMain.Tiles(x).BackColor
                            fMain.AIPlayer4Tile(i).ToolTipText = fMain.Tiles(x).Index
                            fMain.Tiles(x).Visible = False
                            P4Ctr = P4Ctr + 1
                        End If
                    End If
                Next i
                If ValidateWinner = True Then
                    If P4Ctr = 2 Then
                        'MsgBox "Determining the winner...", vbOKOnly + vbInformation, "Player 4's Alert"
                        fMain.MoveStatus.AddItem "Determining the winner..."
                        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
                        Call Game_Winner
                        Exit Function
                    End If
                End If
            Loop
        End If
    'End If
End Function

Function Game_Winner()
    Dim xP1TileValue As Integer, xP2TileValue As Integer, xP3TileValue As Integer, xP4TileValue As Integer
    Dim Play1 As Integer, Play2 As Integer, Play3 As Integer, Play4 As Integer
    Dim GameValArray(3) As Variant
    For i = 0 To 27
        xP1TileValue = xP1TileValue + Val(fMain.AIPlayer1Tile(i).Caption)
        xP2TileValue = xP2TileValue + Val(fMain.AIPlayer2Tile(i).Caption)
        xP3TileValue = xP3TileValue + Val(fMain.AIPlayer3Tile(i).Caption)
        xP4TileValue = xP4TileValue + Val(fMain.AIPlayer4Tile(i).Caption)
    Next i
    GameValArray(0) = xP1TileValue: GameValArray(1) = xP2TileValue: GameValArray(2) = xP3TileValue: GameValArray(3) = xP4TileValue
    Play1 = GameValArray(0): Play2 = GameValArray(1): Play3 = GameValArray(2): Play4 = GameValArray(3)
    'For x = 0 To 3
    '    MsgBox GameValArray(x), vbOKOnly, "Not Sorted"
    'Next x
    Call BubbleSortVariantArray(GameValArray(), True)
    'For x = 0 To 3
    '    MsgBox GameValArray(x), vbOKOnly, "Sorted"
    'Next x
    If Play1 = GameValArray(3) Then
        fMain.MoveStatus.AddItem "Current Winner is: Player 1"
        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
        MsgBox "Current Winner is: Player 1" & vbCrLf & _
               "Current Winner Score is: " & Play1, vbOKOnly + vbInformation, "Game Winner"
    ElseIf Play2 = GameValArray(3) Then
        fMain.MoveStatus.AddItem "Current Winner is: Player 2"
        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
        MsgBox "Current Winner is: Player 2" & vbCrLf & _
               "Current Winner Score is: " & Play2, vbOKOnly + vbInformation, "Game Winner"
    ElseIf Play3 = GameValArray(3) Then
        fMain.MoveStatus.AddItem "Current Winner is: Player 3"
        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
        MsgBox "Current Winner is: Player 3" & vbCrLf & _
               "Current Winner Score is: " & Play3, vbOKOnly + vbInformation, "Game Winner"
    ElseIf Play4 = GameValArray(3) Then
        fMain.MoveStatus.AddItem "Current Winner is: Player 4"
        fMain.MoveStatus.ListIndex = fMain.MoveStatus.NewIndex
        MsgBox "Current Winner is: " & Trim(xCurrent_WinnerName) & vbCrLf & _
               Trim(xCurrent_WinnerName) & " Score is: " & Play4, vbOKOnly + vbInformation, "Game Winner"
        xCurrent_WinnerScore = Play4
        Call fHighestScore.UpdateDatabase
    End If
    GameReset = True: Call fMain.Initialize_Variables: fMain.mnuNewGame.Enabled = True
    fMain.Rack.Caption = "Player 4 Rack"
    fMain.Penalty.Caption = "Player 4 Penalty"
End Function
