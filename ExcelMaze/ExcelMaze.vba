
Sub CreateMaze(ByRef Maze() As Byte, ByRef H As Integer, ByRef W As Integer)
    Dim MH As Integer, MW As Integer, S As Collection, DA As Variant
    Dim I As Integer, D As Long, DI As Integer, Y As Integer, X As Integer
    Dim DY As Integer, DX As Integer, NY As Integer, NX As Integer
    DA = Array(24, 30, 39, 45, 54, 57, 75, 78, 99, 108, 114, 120, _
            135, 141, 147, 156, 177, 180, 198, 201, 210, 216, 225, 228)
    H = WorksheetFunction.Max(1, WorksheetFunction.Min(100, H)): MH = 1 + H + H
    W = WorksheetFunction.Max(1, WorksheetFunction.Min(100, W)): MW = 1 + W + W
    ReDim Maze(MH * MW - 1) As Byte: For I = MH * MW - 1 To 0 Step -1: Maze(I) = 1: Next I
    Set S = New Collection: S.Add H Or 1: S.Add W Or 1
    Do While S.Count > 0
        X = S.Item(S.Count): S.Remove S.Count: Y = S.Item(S.Count): S.Remove S.Count
        Maze(MW * Y + X) = 0
        D = DA(WorksheetFunction.RandBetween(0, 23))
        For I = 0 To 3
            DI = D And 3: D = D \ 4: DY = 0: DX = 0
            If (DI And 2) = 0 Then DY = DI + DI - 1 Else DX = DI + DI - 5
            NY = Y + DY + DY: NX = X + DX + DX
            If (NY > 0) And (NX > 0) And (NY < MH) And (NX < MW) Then
                If Maze(MW * NY + NX) Then
                    Maze(MW * (Y + DY) + (X + DX)) = 0
                    S.Add Y: S.Add X: S.Add NY: S.Add NX
                    Exit For
                End If
            End If
        Next I
    Loop
End Sub

Sub CopyMazeToSheet(ByRef Sheet As Worksheet, ByRef Maze() As Byte, H As Integer, _
        W As Integer, Y0 As Integer, X0 As Integer, WC As Variant, EC As Variant)
    Dim MH As Integer, MW As Integer, I As Integer, Y As Integer, X As Integer, C As Variant
    MH = 1 + H + H: MW = 1 + W + W: I = 0
    For Y = 0 To MH - 1
        For X = 0 To MW - 1
            If Maze(I) Then C = WC Else C = EC
            Sheet.Cells(Y0 + Y, X0 + X).Interior.Color = C
            I = I + 1
        Next X
    Next Y
End Sub

Sub Start()
    Dim CH As Single, CW As Single, CCW As Variant, CCE As Variant
    Dim MH As Integer, MW As Integer, S As Worksheet, M() As Byte
    CH = Sheets("Settings").Evaluate("WALL").RowHeight
    CW = Sheets("Settings").Evaluate("WALL").ColumnWidth
    CCW = Sheets("Settings").Evaluate("WALL").Interior.Color
    CCE = Sheets("Settings").Evaluate("EMPTY").Interior.Color
    MH = Sheets("Settings").Evaluate("HEIGHT").Value
    MW = Sheets("Settings").Evaluate("WIDTH").Value
    Call CreateMaze(M, MH, MW)
    Set S = Sheets.Add: S.Rows.RowHeight = CH: S.Columns.ColumnWidth = CW
    Call CopyMazeToSheet(S, M, MH, MW, 1, 1, CCW, CCE)
End Sub
