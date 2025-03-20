Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
    Static firstSelection As Range
    Dim ws As Worksheet
    Set ws = Me
    
    ' Check turn and if selection is within A1:H8
    If ws.Range("J1").Value <> "White" Then Exit Sub
    If Not Intersect(Target, ws.Range("A1:H8")) Is Nothing And Target.Count = 1 Then
        If firstSelection Is Nothing Then
            ' First click: Select White piece
            If IsWhitePiece(ws.Cells(Target.Row, Target.Column).Value) Then
                Set firstSelection = Target
                firstSelection.Interior.Color = RGB(255, 255, 0) ' Highlight
            End If
        Else
            ' Second click: Attempt move
            If IsLegalMove(firstSelection, Target) Then
                Target.Value = firstSelection.Value
                firstSelection.Value = ""
                firstSelection.Interior.ColorIndex = xlNone
                ws.Range("J1").Value = "Black"
                Application.OnTime Now + TimeValue("00:00:01"), "ComputerMove"
            Else
                MsgBox "Illegal move!"
                firstSelection.Interior.ColorIndex = xlNone
            End If
            Set firstSelection = Nothing
        End If
    End If
End Sub

Private Function IsWhitePiece(piece As String) As Boolean
    Dim pieceCode As Long
    If piece <> "" Then
        pieceCode = AscW(piece)
        IsWhitePiece = (pieceCode >= &H2654 And pieceCode <= &H2659)
    End If
End Function

Private Function IsLegalMove(startCell As Range, endCell As Range) As Boolean
    Dim ws As Worksheet
    Set ws = startCell.Parent
    Dim piece As String
    piece = startCell.Value
    Dim startRow As Integer, startCol As Integer
    Dim endRow As Integer, endCol As Integer
    startRow = startCell.Row
    startCol = startCell.Column
    endRow = endCell.Row
    endCol = endCell.Column
    
    ' No move if target has own piece or no change
    If IsWhitePiece(endCell.Value) Or (startRow = endRow And startCol = endCol) Then
        Exit Function
    End If
    
    Select Case piece
        Case ChrW(&H2659) ' White Pawn
            If startCol = endCol Then ' Same column
                If startRow = 7 And endRow = 5 And ws.Cells(6, startCol).Value = "" And endCell.Value = "" Then
                    IsLegalMove = True ' Two-square move from start
                ElseIf endRow = startRow - 1 And endCell.Value = "" Then
                    IsLegalMove = True ' One-square move
                End If
            End If
        Case ChrW(&H2656) ' White Rook
            If startRow = endRow Then ' Horizontal
                IsLegalMove = IsPathClear(ws, startRow, startCol, endRow, endCol, True)
            ElseIf startCol = endCol Then ' Vertical
                IsLegalMove = IsPathClear(ws, startRow, startCol, endRow, endCol, False)
            End If
        Case ChrW(&H2658) ' White Knight
            Dim rowDiff As Integer, colDiff As Integer
            rowDiff = Abs(endRow - startRow)
            colDiff = Abs(endCol - startCol)
            If (rowDiff = 2 And colDiff = 1) Or (rowDiff = 1 And colDiff = 2) Then
                IsLegalMove = True
            End If
    End Select
End Function

Private Function IsPathClear(ws As Worksheet, startRow As Integer, startCol As Integer, _
                            endRow As Integer, endCol As Integer, isHorizontal As Boolean) As Boolean
    Dim i As Integer
    IsPathClear = True
    If isHorizontal Then
        Dim minCol As Integer, maxCol As Integer
        minCol = IIf(startCol < endCol, startCol + 1, endCol + 1)
        maxCol = IIf(startCol < endCol, endCol - 1, startCol - 1)
        For i = minCol To maxCol
            If ws.Cells(startRow, i).Value <> "" Then
                IsPathClear = False
                Exit Function
            End If
        Next i
    Else
        Dim minRow As Integer, maxRow As Integer
        minRow = IIf(startRow < endRow, startRow + 1, endRow + 1)
        maxRow = IIf(startRow < endRow, endRow - 1, startRow - 1)
        For i = minRow To maxRow
            If ws.Cells(i, startCol).Value <> "" Then
                IsPathClear = False
                Exit Function
            End If
        Next i
    End If
End Function

