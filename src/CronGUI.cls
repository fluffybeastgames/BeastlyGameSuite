VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CronGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wsCron As Worksheet
Public rngBoardTopLeft As Range
Private rngBoard As Range
Private rngDebugStats As Range
Private arrBoard() As Integer

'Set rngBoardTopLeft = wsCron.Range("Game_Board_Top_Left")

Sub init(ByVal rows As Integer, ByVal cols As Integer)
    ReDim arrBoard(rows - 1, cols - 1)
    
    Set wsCron = ThisWorkbook.Worksheets("Cron")
    Set rngBoardTopLeft = wsCron.Range("Game_Board_Top_Left")
    Set rngBoard = wsCron.Range(rngBoardTopLeft.Address, rngBoardTopLeft.Offset(rows - 1, cols - 1).Address)
    Set rngDebugStats = wsCron.Range("O1:Z5")
    
    prep_game_sheet
End Sub


Sub prep_game_sheet()
    ' Result colors and borders/text
    With wsCron.Cells
        .Interior.Color = RGB(55, 55, 55)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlRight ' make it easier to read debug text
        .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    End With
    
    'Setup board area
    With rngBoard
        .Interior.Color = RGB(68, 84, 106)
        .ColumnWidth = 0.5 ' 3 = 26 pixels on my device
        .RowHeight = 5 ' 24 = 32 pixels on my device
        .HorizontalAlignment = xlCenter ' center text this range
        .VerticalAlignment = xlCenter ' center text this range
        .Font.Bold = True
            
        'Add thick borders around the edges of the board
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        

    End With
            
    rngDebugStats.Cells(1, 1).value = "Elapsed"
    rngDebugStats.Cells(1, 2).value = "Outer Loop"
    rngDebugStats.Cells(1, 3).value = "Ticks"
    rngDebugStats.Cells(1, 4).value = "Frames"
    rngDebugStats.Cells(1, 5).value = "FPS"
    rngDebugStats.Cells(1, 6).value = "Score"
    
    '' Hide rows we won't need
    'wsCron.Range("A20:A" & rows.Count).EntireRow.Hidden = True
    'wsCron.Range("A20:A" & rows.Count).EntireRow.Hidden = True
    'wsCron.Range(Cells(1, 20), Cells(1, Columns.Count)).EntireColumn.Hidden = True
End Sub

'gui.update_debug_stats outer_loop_no, ti, frame_no, game_timer

Sub update_debug_stats(outer_loop_no, tick_no, frame_no, game_timer, score)

    game_duration = Timer - game_timer
    rngDebugStats.Cells(2, 1).value = Round(game_duration, 2)
    rngDebugStats.Cells(2, 2).value = outer_loop_no
    rngDebugStats.Cells(2, 3).value = tick_no
    rngDebugStats.Cells(2, 4).value = frame_no
    rngDebugStats.Cells(2, 5).value = Round(frame_no / game_duration, 2)
    rngDebugStats.Cells(2, 6).value = score
    
        
End Sub


Public Sub render_board(game As CronGame)
    Dim str_rng_grey As String, str_rng_snek As String, str_rng_apples As String
    
    For r = 0 To UBound(arrBoard, 1)
        For c = 0 To UBound(arrBoard, 2)
            arrBoard(r, c) = 0
        
        Next c
    Next r
    
    Dim i As Long
    For i = 1 To game.cron_body_player_1.Count
        arrBoard(game.cron_body_player_1(i).row, game.cron_body_player_1(i).col) = 1
    Next i
    
    For i = 1 To game.cron_body_player_2.Count
        arrBoard(game.cron_body_player_2(i).row, game.cron_body_player_2(i).col) = 2
    Next i
    
    For r = 0 To UBound(arrBoard, 1)
        For c = 0 To UBound(arrBoard, 2)
            If arrBoard(r, c) = 1 Then
                rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(32, 105, 224)
            ElseIf arrBoard(r, c) = 2 Then
                rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(247, 111, 237)
            Else
                If rngBoard(1, 1).Offset(r, c).Interior.Color <> RGB(68, 84, 106) Then
                    rngBoard(1, 1).Offset(r, c).Interior.Color = RGB(68, 84, 106)
                End If
            End If
        Next c
    Next r
    
        
End Sub