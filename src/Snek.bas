Attribute VB_Name = "Snek"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Sub create_game_sheet()
    ' First see if it exists.. if so prompt user if they want to delete it
        
    Dim boolAlreadyExists As Boolean
    boolAlreadyExists = False
    
    For Each s In ThisWorkbook.Sheets
        If s.name = "Snek" Then
            MsgBox "There is already a sheet named 'Snek'. Please remove or rename it."
            boolAlreadyExists = True
        End If
    Next s
    
    If boolAlreadyExists Then
        Debug.Print "Not creating"
    Else
        Debug.Print "Creating"
        Application.ScreenUpdating = False
        ' On Error GoTo errCleanExit
        
        'Create new sheet, name it 'Snek', add game board, start button, high score, and maybe a button w/ a user form for game settings
        
        Dim wsSnek As Worksheet, rngBoard As Range
        Dim rngBtnNewGame As Range, rngBtnGameSettings As Range, btnNewGame As Button, btnGameSettings As Button
            
        Set wsSnek = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)) 'Create the new worksheet
        wsSnek.name = "Snek"
        
        With wsSnek.Range("C3")
            .Font.Bold = True
            .Font.Size = 24
            .value = "SNEK"
        End With
        
        Set rngBoard = wsSnek.Range("B5:M12")
        wsSnek.Names.Add name:="Game_Board", RefersTo:=rngBoard
        
                
        'Add thin borders within the board and thick borders around it
        With rngBoard.Borders
            .LineStyle = xlContinuous
            .Color = vbGrey
            .Weight = xlThin
        End With
        
        rngBoard.Borders(xlEdgeTop).Weight = xlThick
        rngBoard.Borders(xlEdgeBottom).Weight = xlThick
        rngBoard.Borders(xlEdgeLeft).Weight = xlThick
        rngBoard.Borders(xlEdgeRight).Weight = xlThick
        
        With rngBoard
            .ColumnWidth = 3.5 ' 3 = 26 pixels on my device
            .RowHeight = 25 ' 24 = 32 pixels on my device
            .HorizontalAlignment = xlCenter ' center text this range
            .VerticalAlignment = xlCenter ' center text this range
            .Font.Bold = True
        End With
        
        
        wsSnek.Range("A:AA").Interior.Color = RGB(32, 32, 32)
        rngBoard.Interior.Color = RGB(200, 200, 200)
        
        wsSnek.Range("1:19").Font.Color = RGB(255, 255, 255)
        
        wsSnek.Range("15:19").HorizontalAlignment = xlRight ' make it easier to read debug text
        
        ' Hide rows we won't need
        wsSnek.Range("A20:A" & Rows.Count).EntireRow.Hidden = True
        wsSnek.Range("A20:A" & Rows.Count).EntireRow.Hidden = True
        wsSnek.Range(Cells(1, 20), Cells(1, Columns.Count)).EntireColumn.Hidden = True

        'rng.EntireRow.Hidden = True
        
        
        'Add New and Settings buttons and bind functionality to them
        
        Set rngBtnNewGame = wsSnek.Range("O6")
        Set rngBtnGameSettings = wsSnek.Range("O10")
                     
        Set btnNewGame = wsSnek.Buttons.Add(rngBtnNewGame.Left, rngBtnNewGame.Top, rngBtnNewGame.Width, rngBtnNewGame.Height)
        Set btnGameSettings = wsSnek.Buttons.Add(rngBtnGameSettings.Left, rngBtnGameSettings.Top, rngBtnGameSettings.Width, rngBtnGameSettings.Height)
        
        btnNewGame.Caption = "New Game"
        btnNewGame.name = "BtnNewGame"
        btnNewGame.OnAction = "start_new_game"
        
        btnGameSettings.Caption = "Settings"
        btnGameSettings.name = "BtnSettings"
        btnGameSettings.OnAction = "open_game_settings"
        
             
        'Clean up
        Set wsSnek = Nothing
        Set rngBoard = Nothing
        Set rngBtnNewGame = Nothing
        Set rngBtnGameSettings = Nothing
        Set btnNewGame = Nothing
        Set btnGameSettings = Nothing
        
        Application.ScreenUpdating = True
        
    
        
    End If
    
    Exit Sub
    
errCleanExit: ' In case of error, make sure we turn screen updating back on - otherwise the user will be sure to get annoyed!
    Application.ScreenUpdating = True
    
    MsgBox "Encountered an error:" & vbNewLine & Err.Number, Err.Description
End Sub


Sub start_new_game()
    game_loop
        
End Sub

Sub open_game_settings()
    Debug.Print "open_game_settings"
    Debug.Print "TODO"
End Sub

Private Function is_valid_move(cur_row, cur_col, move_direction, arrBoard, snek_body)
    'Debug.Print cur_row & vbTab & cur_col & vbTab & move_direction
    If move_direction = "N" And cur_row > 0 Then
       is_valid_move = arrBoard(cur_row - 1, cur_col) <> 1
        
    ElseIf move_direction = "S" And cur_row < UBound(arrBoard, 1) - 1 Then
        is_valid_move = True
        is_valid_move = arrBoard(cur_row + 1, cur_col) <> 1
    
    ElseIf move_direction = "W" And cur_col > 0 Then
        is_valid_move = True
        is_valid_move = arrBoard(cur_row, cur_col - 1) <> 1
    
    ElseIf move_direction = "E" And cur_col < UBound(arrBoard, 2) - 1 Then
        is_valid_move = True
        is_valid_move = arrBoard(cur_row, cur_col + 1) <> 1
        
    Else
        is_valid_move = False
    End If
End Function

Sub game_loop()
    Debug.Print "Initiating game loop"
    
    Dim wsSnek As Worksheet: Set wsSnek = ThisWorkbook.Worksheets("Snek")
    Dim rngBoard As Range: Set rngBoard = wsSnek.Range("Game_Board")
    
    Dim snek_body As New Collection
    Dim apples As New Collection
    
    
    Dim game_loop_no As Long, game_loop_inner_no As Long ' Number of iterations of the outer do while game_on loop, and number of times spent doing DoingEvents
    Dim loop_duration As Double: loop_duration = 1 / 1000
    
    Dim cycle_no As Long ' How many game steps have we taken, aka how many times have we gone through the outer Do While loop
    Dim cycle_timer As Double
    Dim cycle_duration As Double: cycle_duration = 1 / 2 ' how long between each turn (ie move de snek)
    
    Dim frame_no As Long ' How many frames have been rendered
    Dim frame_timer As Double ' A timestamp to let us know if enough time has passed to render a new frame
    Dim frame_duration As Double: frame_duration = 1 / 60 ' preferred frames (in SFP not FPS)
    Dim frame_lock_duration: frame_lock_duration = frame_duration * 0.9 ' if the duration exceeds this,
    Dim move_direction As String: move_direction = "E"
    
    Dim game_on As Boolean
    game_on = True
    game_timer = Timer ' Records the entire length of the game
    
    loop_timer = Timer ' How long before we see if it's cycle time? DoEvents in here instead
    cycle_timer = Timer ' Records the length since the last game step.. once it passes the threshold of cycle_duration, the game advances one step
    frame_timer = Timer ' Records the length since the last render.. once it passes the threshold of frame_duration, the game advances redraws the current state
    
    
    rngBoard.Interior.Color = RGB(200, 200, 200) 'Remove the old game state (if any)
    wsSnek.Range("P15:S19").ClearContents ' get rid of any old debug dialog before we get started
    


    DisableArrowKeys
    
    Dim num_rows As Integer, num_cols As Integer
    num_rows = 8
    num_cols = 12
    Dim arrBoard() As Integer
    ReDim arrBoard(num_rows, num_cols)
    
    Dim cur_row As Integer, cur_col As Integer
    cur_row = 1
    cur_col = 1
    arrBoard(cur_row, cur_col) = 1
    
    Dim s As SnekSegment: Set s = New SnekSegment
    s.row = cur_row
    s.col = cur_col
    snek_body.Add s
    
    
    add_apple arrBoard, rngBoard, apples
    'add_apple arrBoard, rngBoard, apples
    'add_apple arrBoard, rngBoard, apples
    
    
    'render_board arrBoard, rngBoard
    render_board rngBoard, snek_body
    
    Do While game_on
        game_loop_no = game_loop_no + 1
        'wsSnek.Range("P18").value = "Loop"
        'wsSnek.Range("P19").value = game_loop_no
        
        move_direction = check_for_keyboard_input(move_direction)

        
        
        Do Until Timer - loop_timer >= loop_duration ' Wait until ready to process cycle
            game_loop_inner_no = game_loop_inner_no + 1
            
            DoEvents
            
        Loop
        loop_timer = Timer
            
           
        If Timer - cycle_timer >= cycle_duration Then ' Handle user input and update the simulation by one step
            cycle_no = cycle_no + 1
            cycle_timer = Timer
            
            If is_valid_move(cur_row, cur_col, move_direction, arrBoard, snek_body) Then
                
                If move_direction = "N" Then
                    cur_row = cur_row - 1
                ElseIf move_direction = "S" Then
                    cur_row = cur_row + 1
                ElseIf move_direction = "W" Then
                    cur_col = cur_col - 1
                ElseIf move_direction = "E" Then
                    cur_col = cur_col + 1
                End If
            Else
                game_on = False
                Debug.Print "GAME OVER"
            End If
            
            ' Debug.Print "New address " & cur_row & ", " & cur_col
            arrBoard(cur_row, cur_col) = 1
            
            Set s = New SnekSegment
            s.row = cur_row
            s.col = cur_col
            snek_body.Add s
            
            apple_in_new_spot = False
            For i = 1 To apples.Count:
                Debug.Print "apple " & apples(i).row & apples(i).col
                
                If apples(i).row = s.row And apples(i).col = s.col Then
                    Debug.Print "apple found"
                    apple_in_new_spot = True
                    apples.Remove i
                    Exit For
                End If
            Next i
            
            If apple_in_new_spot Then
                
                add_apple arrBoard, rngBoard, apples
            Else
                
                'Remove the trailing end
                rngBoard(1, 1).Offset(snek_body(1).row, snek_body(1).col).Interior.Color = RGB(200, 200, 200)
                arrBoard(snek_body(1).row, snek_body(1).col) = 0
                snek_body.Remove 1
                'remove_snek_seg rngBoard, snek_body
            End If
            
            render_board rngBoard, snek_body
            
            'Increase difficulty as the snek gets longer
            If snek_body.Count > 10 Then
                cycle_duration = 1 / 8
            ElseIf snek_body.Count > 6 Then
                cycle_duration = 1 / 4
            ElseIf snek_body.Count > 3 Then
                cycle_duration = 1 / 3
            Else
                cycle_duration = 1 / 2
            End If
                
            
            'Update performance stats text
            wsSnek.Range("R18").value = "Cycle"
            wsSnek.Range("R19").value = cycle_no
            wsSnek.Range("Q18").value = "Loop Inner"
            wsSnek.Range("Q19").value = game_loop_no + game_loop_inner_no ' the number of game loops plus any pauses within
            
            game_duration = Timer - game_timer
            wsSnek.Range("P15").value = "Time Elapsed"
            wsSnek.Range("P16").value = game_duration
            
            wsSnek.Range("Q15").value = "Score"
            wsSnek.Range("Q16").value = snek_body.Count
            
            wsSnek.Range("R15").value = "FPS"
            If game_duration > 0 Then
                wsSnek.Range("R16").value = frame_no / game_duration
            End If
            
        End If
                       
        If Timer - frame_timer >= frame_duration Then
            frame_no = frame_no + 1
            frame_timer = Timer
            
            'update_debug_display frame_no, game_loop_no
            wsSnek.Range("S18").value = "Frames"
            wsSnek.Range("S19").value = frame_no
            
        
            wsSnek.Range("P18").value = "Loop"
            wsSnek.Range("P19").value = game_loop_no
        
''''Initial test showed no drop (57 FPS --> 57 FPS out of a goal of 60) but we're not drawing or computing much yet, so let's wait and see..
'        ElseIf Timer - frame_timer >= frame_lock_duration Then
'            Do Until Timer - frame_timer >= frame_duration ' Wait until ready to process cycle
'                game_loop_inner_no = game_loop_inner_no + 1
'                DoEvents
'            Loop
        
        
        
        End If
        
        
        ' End game loop check
       ' If frame_no >= 1080 Then
        '    game_on = False
      '  End If
        

        
    Loop
    
    game_duration = Timer - game_timer
    Debug.Print "Final Values: " & frame_no & " frames in " & cycle_no & " cycles over " & game_duration & " seconds, resulting in " & frame_no / game_duration & " fps"
    wsSnek.Range("P15").value = "Time Elapsed"
    wsSnek.Range("P16").value = game_duration

    wsSnek.Range("R15").value = "FPS"
    wsSnek.Range("R16").value = frame_no / game_duration
    
    EnableArrowKeys
    
    Set wsSnek = Nothing
    Set rngBoard = Nothing

End Sub

Private Sub add_apple(arrBoard, rngBoard, apples)
    Dim row As Integer, col As Integer, a As SnekApple
    empty_spot_found = False
    
    Do While Not empty_spot_found
        row = Rnd * (UBound(arrBoard, 1) - 1)
        col = Rnd * (UBound(arrBoard, 2) - 1)
        
'        Debug.Print "rnd " & row, col
        
        If arrBoard(row, col) < 1 Then
            empty_spot_found = True
            arrBoard(row, col) = 2
            rngBoard(1, 1).Offset(row, col).Interior.Color = RGB(20, 200, 20)
            
    
            Set a = New SnekApple
            a.row = row
            a.col = col
            apples.Add a
        End If
    Loop
    
    


End Sub
Private Function check_for_keyboard_input(move_direction)
        If GetAsyncKeyState(vbKeyUp) And move_direction <> "S" Then ' Don't check this during the cycle loop as it will only catch it if they key is held down when it fires
            check_for_keyboard_input = "N"
        ElseIf GetAsyncKeyState(vbKeyDown) And move_direction <> "N" Then
            check_for_keyboard_input = "S"
        ElseIf GetAsyncKeyState(vbKeyLeft) And move_direction <> "E" Then
            check_for_keyboard_input = "W"
        ElseIf GetAsyncKeyState(vbKeyRight) And move_direction <> "W" Then '
            check_for_keyboard_input = "E"
        Else
            check_for_keyboard_input = move_direction
        End If
End Function

Private Sub render_board(rngBoard, snek_body)
    
    Dim i As Long
    For i = 1 To snek_body.Count
        rngBoard(1, 1).Offset(snek_body(i).row, snek_body(i).col).Interior.Color = RGB(200, 20, 20)
        
    Next i
    
End Sub


Private Sub render_board_old(arrBoard, rngBoard)
    
    For i = 0 To UBound(arrBoard, 1)
        For j = 0 To UBound(arrBoard, 2)
            If arrBoard(i, j) = 1 Then
                rngBoard(1, 1).Offset(i, j).Interior.Color = RGB(200, 20, 20)
            ElseIf arrBoard(i, j) = -1 Then
                rngBoard(1, 1).Offset(i, j).Interior.Color = RGB(200, 200, 200)
            End If
        Next j
    Next i
    
End Sub

Sub DisableArrowKeys()
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
End Sub
Sub EnableArrowKeys()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
End Sub




Sub coll_test():
    ' Create collection
    Dim snake_cells As New Collection
    
    
    ' Read 100 values to collection
    Dim c As Range
    For Each c In Sheet1.Range("A1:A100")
        ' This line is used to add all the values
        collMarks.Add c.value
    Next
    

End Sub

