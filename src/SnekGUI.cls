VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnekGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wsSnek As Worksheet
Public rngBoardTopLeft As Range
Private rngBoard As Range
Private rngDebugStats As Range
Private arrBoard() As Integer

'Set rngBoardTopLeft = wsSnek.Range("Game_Board_Top_Left")

Sub init(ByVal rows As Integer, ByVal cols As Integer)
    ReDim arrBoard(rows - 1, cols - 1)
    
    Set wsSnek = ThisWorkbook.Worksheets("Snek")
    Set rngBoardTopLeft = wsSnek.Range("Game_Board_Top_Left")
    Set rngBoard = wsSnek.Range(rngBoardTopLeft.Address, rngBoardTopLeft.Offset(rows - 1, cols - 1).Address)
    Set rngDebugStats = wsSnek.Range("O1:Z5")
    
    prep_game_sheet
    
End Sub


Sub prep_game_sheet()
    Application.ScreenUpdating = False

    
    ' Result colors and borders/text
    With wsSnek.Cells
        .Interior.Color = RGB(110, 170, 70)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlRight ' make it easier to read debug text
        .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    End With
    
    'Setup board area
    With rngBoard
        .Interior.Color = RGB(200, 200, 200)
        .ColumnWidth = 3.5 ' 3 = 26 pixels on my device
        .RowHeight = 25 ' 24 = 32 pixels on my device
        .HorizontalAlignment = xlCenter ' center text this range
        .VerticalAlignment = xlCenter ' center text this range
        .Font.Bold = True
            
        'Add thin borders within the board and thicker borders around it
        With .Borders
            .LineStyle = xlContinuous
            .Color = vbGrey
            .Weight = xlThin
        End With
    
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
                
    rngDebugStats.Cells(1, 1).Value = "Elapsed"
    rngDebugStats.Cells(1, 2).Value = "Outer Loop"
    rngDebugStats.Cells(1, 3).Value = "Ticks"
    rngDebugStats.Cells(1, 4).Value = "Frames"
    rngDebugStats.Cells(1, 5).Value = "FPS"
    rngDebugStats.Cells(1, 6).Value = "Score"
    
    

    With wsSnek.Range("B3")
        .Font.Bold = True
        .Font.Size = 24
        .Value = "SNEK"
        .HorizontalAlignment = xlLeft
    End With
    With wsSnek.Range("B4")
        .Font.Bold = True
        .Font.Size = 14
        .Value = "Feed the hungry critter! Change directions with the arrow keys"
        .HorizontalAlignment = xlLeft
    End With

    
    '' Hide rows we won't need
    'wsSnek.Range("A20:A" & rows.Count).EntireRow.Hidden = True
    'wsSnek.Range("A20:A" & rows.Count).EntireRow.Hidden = True
    'wsSnek.Range(Cells(1, 20), Cells(1, Columns.Count)).EntireColumn.Hidden = True
    
    Application.ScreenUpdating = True
End Sub

'gui.update_debug_stats outer_loop_no, ti, frame_no, game_timer
Sub update_debug_stats(outer_loop_no, tick_no, frame_no, game_timer, score)

    game_duration = Timer - game_timer
    rngDebugStats.Cells(2, 1).Value = Round(game_duration, 2)
    rngDebugStats.Cells(2, 2).Value = outer_loop_no
    rngDebugStats.Cells(2, 3).Value = tick_no
    rngDebugStats.Cells(2, 4).Value = frame_no
    rngDebugStats.Cells(2, 5).Value = Round(frame_no / game_duration, 2)
    rngDebugStats.Cells(2, 6).Value = score
    
        
End Sub

Public Sub render_board(game As SnekGame)

    Dim str_rng_grey As String, str_rng_snek As String, str_rng_apples As String
    
    For r = 0 To UBound(arrBoard, 1)
        For c = 0 To UBound(arrBoard, 2)
            arrBoard(r, c) = 0
        
        Next c
    Next r
    
    Dim i As Long
    For i = 1 To game.snek_body.Count
        arrBoard(game.snek_body(i).row, game.snek_body(i).col) = 1
    Next i
    
    For i = 1 To game.apples.Count
        arrBoard(game.apples(i).row, game.apples(i).col) = 2
    Next i
    
    For r = 0 To UBound(arrBoard, 1)
        For c = 0 To UBound(arrBoard, 2)
            If arrBoard(r, c) = 1 Then
                rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(200, 20, 20)
            ElseIf arrBoard(r, c) = 2 Then
                rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(20, 200, 20)
            Else
                If rngBoard(1, 1).Offset(r, c).Interior.Color <> RGB(200, 200, 200) Then
                    rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(200, 200, 200)
                End If
            End If
        Next c
    Next r
     
End Sub


