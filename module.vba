Sub SetupChessboard()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If MsgBox("Clear the sheet to set up the chessboard?", vbYesNo) = vbNo Then Exit Sub
    ws.Cells.Clear
    
    ws.Range("A1:H8").ColumnWidth = 4
    ws.Range("A1:H8").RowHeight = 30
    
    Dim i As Integer, j As Integer
    For i = 1 To 8
        For j = 1 To 8
            With ws.Cells(i, j)
                If (i + j) Mod 2 = 0 Then
                    .Interior.Color = RGB(240, 217, 181)
                Else
                    .Interior.Color = RGB(181, 136, 99)
                End If
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 20
            End With
        Next j
    Next i
    
    ws.Cells(8, 1).Value = ChrW(&H2656) ' White Rook
    ws.Cells(8, 2).Value = ChrW(&H2658) ' White Knight
    ws.Cells(8, 3).Value = ChrW(&H2657) ' White Bishop
    ws.Cells(8, 4).Value = ChrW(&H2655) ' White Queen
    ws.Cells(8, 5).Value = ChrW(&H2654) ' White King
    ws.Cells(8, 6).Value = ChrW(&H2657) ' White Bishop
    ws.Cells(8, 7).Value = ChrW(&H2658) ' White Knight
    ws.Cells(8, 8).Value = ChrW(&H2656) ' White Rook
    For j = 1 To 8
        ws.Cells(7, j).Value = ChrW(&H2659) ' White Pawns
    Next j
    
    ws.Cells(1, 1).Value = ChrW(&H265C) ' Black Rook
    ws.Cells(1, 2).Value = ChrW(&H265E) ' Black Knight
    ws.Cells(1, 3).Value = ChrW(&H265D) ' Black Bishop
    ws.Cells(1, 4).Value = ChrW(&H265B) ' Black Queen
    ws.Cells(1, 5).Value = ChrW(&H265A) ' Black King
    ws.Cells(1, 6).Value = ChrW(&H265D) ' Black Bishop
    ws.Cells(1, 7).Value = ChrW(&H265E) ' Black Knight
    ws.Cells(1, 8).Value = ChrW(&H265C) ' Black Rook
    For j = 1 To 8
        ws.Cells(2, j).Value = ChrW(&H265F) ' Black Pawns
    Next j
    
    For i = 1 To 8
        ws.Cells(i, 9).Value = 9 - i
        ws.Cells(9, i).Value = Chr(64 + i)
    Next i
    
    ws.Range("J1").Value = "White"
End Sub

Sub ComputerMove()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim blackPieces As Collection
    Set blackPieces = New Collection
    Dim i As Integer, j As Integer
    
    ' Collect Black pieces
    For i = 1 To 8
        For j = 1 To 8
            If IsBlackPiece(ws.Cells(i, j).Value) Then
                blackPieces.Add ws.Cells(i, j)
            End If
        Next j
    Next i
    
    ' Try moves until a legal one is found
    Randomize
    Dim pieceCell As Range, targetCell As Range
    Dim attempts As Integer
    attempts = 0
    Do While attempts < 100 ' Prevent infinite loop
        Set pieceCell = blackPieces(Int(Rnd * blackPieces.Count) + 1)
        Set targetCell = GetLegalMoveForPiece(ws, pieceCell)
        If Not targetCell Is Nothing Then
            targetCell.Value = pieceCell.Value
            pieceCell.Value = ""
            Exit Do
        End If
        attempts = attempts + 1
    Loop
    
    ws.Range("J1").Value = "White"
End Sub

Private Function IsBlackPiece(piece As String) As Boolean
    Dim pieceCode As Long
    If piece <> "" Then
        pieceCode = AscW(piece)
        IsBlackPiece = (pieceCode >= &H265A And pieceCode <= &H265F)
    End If
End Function

Private Function GetLegalMoveForPiece(ws As Worksheet, pieceCell As Range) As Range
    Dim piece As String
    piece = pieceCell.Value
    Dim startRow As Integer, startCol As Integer
    startRow = pieceCell.Row
    startCol = pieceCell.Column
    Dim possibleMoves As Collection
    Set possibleMoves = New Collection
    Dim i As Integer, j As Integer
    
    Select Case piece
        Case ChrW(&H265F) ' Black Pawn
            If startRow = 2 And ws.Cells(3, startCol).Value = "" And ws.Cells(4, startCol).Value = "" Then
                possibleMoves.Add ws.Cells(4, startCol)
            End If
            If ws.Cells(startRow + 1, startCol).Value = "" Then
                possibleMoves.Add ws.Cells(startRow + 1, startCol)
            End If
        Case ChrW(&H265C) ' Black Rook
            ' Check vertical up
            For i = startRow - 1 To 1 Step -1
                If ws.Cells(i, startCol).Value = "" Then
                    possibleMoves.Add ws.Cells(i, startCol)
                Else
                    Exit For
                End If
            Next i
            ' Check vertical down
            For i = startRow + 1 To 8
                If ws.Cells(i, startCol).Value = "" Then
                    possibleMoves.Add ws.Cells(i, startCol)
                Else
                    Exit For
                End If
            Next i
            ' Check horizontal left
            For j = startCol - 1 To 1 Step -1
                If ws.Cells(startRow, j).Value = "" Then
                    possibleMoves.Add ws.Cells(startRow, j)
                Else
                    Exit For
                End If
            Next j
            ' Check horizontal right
            For j = startCol + 1 To 8
                If ws.Cells(startRow, j).Value = "" Then
                    possibleMoves.Add ws.Cells(startRow, j)
                Else
                    Exit For
                End If
            Next j
        Case ChrW(&H265E) ' Black Knight
            Dim moves(7, 1) As Integer
            moves(0, 0) = -2: moves(0, 1) = -1
            moves(1, 0) = -2: moves(1, 1) = 1
            moves(2, 0) = -1: moves(2, 1) = -2
            moves(3, 0) = -1: moves(3, 1) = 2
            moves(4, 0) = 1: moves(4, 1) = -2
            moves(5, 0) = 1: moves(5, 1) = 2
            moves(6, 0) = 2: moves(6, 1) = -1
            moves(7, 0) = 2: moves(7, 1) = 1
            For i = 0 To 7
                Dim newRow As Integer, newCol As Integer
                newRow = startRow + moves(i, 0)
                newCol = startCol + moves(i, 1)
                If newRow >= 1 And newRow <= 8 And newCol >= 1 And newCol <= 8 Then
                    If ws.Cells(newRow, newCol).Value = "" Then
                        possibleMoves.Add ws.Cells(newRow, newCol)
                    End If
                End If
            Next i
    End Select
    
    If possibleMoves.Count > 0 Then
        Set GetLegalMoveForPiece = possibleMoves(Int(Rnd * possibleMoves.Count) + 1)
    End If
End Function
