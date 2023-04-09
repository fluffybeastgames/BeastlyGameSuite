VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private outer_loop_duration As Double
Private outer_loop_timer As Double
Private outer_loop_no As Long ' Number of iterations of the outer do while game_on loop


Private tick_timer As Double ' keep track of when the previous tick was initiated
Private tick_no As Long ' How many game steps have we taken, aka how many times have we gone through the outer Do While loop

Private frame_duration As Double:
Private frame_timer As Double ' A timestamp to let us know if enough time has passed to render a new frame
Private frame_no As Long ' How many frames have been rendered

Private game_timer As Double

Private Sub Class_Initialize()
    outer_loop_duration = 1 / 1000 ' target seconds per check if tick or frame needs to be advanced
    frame_duration = 1 / 60 ' target seconds per frame

    game_timer = Timer          ' Records the entire length of the game
    outer_loop_timer = Timer    ' How long before we see if it's tick time? DoEvents in here instead
    tick_timer = Timer          ' Records the length since the last game step.. once it passes the threshold of tick_duration, the game advances one step
    frame_timer = Timer         ' Records the length since the last render.. once it passes the threshold of frame_duration, the game advances redraws the current state
    
    outer_loop_no = 0
    tick_no = 0
    frame_no = 0
End Sub

Sub run_game_loop(game, gui)
    Debug.Print "Starting game loop"
    move_dir_p1 = "S"
    move_dir_p2 = "S"
    
    Do While game.game_on
        outer_loop_no = outer_loop_no + 1
        
        ' Check for user input
        move_dir_p1 = check_for_keyboard_input_arrows(move_dir_p1)
        move_dir_p2 = check_for_keyboard_input_WASD(move_dir_p2)
        
        If Timer - tick_timer >= game.tick_duration Then
            ' Process user input and update the simulation by one step
            tick_no = tick_no + 1
            tick_timer = Timer
            
            'Debug.Print "--Tick " & vbTab & tick_no
            
            If game.num_players = 1 Then
                game.game_on = game.tick(move_dir_p1) ' tick tock don't let the game stop - main game step logic is in here!
            Else
                game.game_on = game.tick(move_dir_p1, move_dir_p2) ' tick tock don't let the game stop - main game step logic is in here!
            End If
            
        End If
        
        If Timer - frame_timer >= frame_duration Then
            frame_no = frame_no + 1
            frame_timer = Timer
            gui.render_board game
            'gui.update_debug_stats outer_loop_no, tick_no, frame_no, game_timer, game.snek_body.Count - 1
            gui.update_debug_stats outer_loop_no, tick_no, frame_no, game_timer, -1
            
        End If
    
        DoEvents ' Yield system if need be
        
        Do Until Timer - outer_loop_timer >= outer_loop_duration ' Wait until ready to process tick
            game_loop_inner_no = game_loop_inner_no + 1
            
        Loop
        outer_loop_timer = Timer
        
    Loop
    
    game.end_game
    
    
    
End Sub


Private Function check_for_keyboard_input_arrows(move_direction)
    If GetAsyncKeyState(vbKeyUp) And move_direction <> "S" Then
        check_for_keyboard_input_arrows = "N"
    ElseIf GetAsyncKeyState(vbKeyDown) And move_direction <> "N" Then
        check_for_keyboard_input_arrows = "S"
    ElseIf GetAsyncKeyState(vbKeyLeft) And move_direction <> "E" Then
        check_for_keyboard_input_arrows = "W"
    ElseIf GetAsyncKeyState(vbKeyRight) And move_direction <> "W" Then '
        check_for_keyboard_input_arrows = "E"
    Else
        check_for_keyboard_input_arrows = move_direction
    End If
End Function



Private Function check_for_keyboard_input_WASD(move_direction)
    If GetAsyncKeyState(VK_W) And move_direction <> "S" Then
        check_for_keyboard_input_WASD = "N"
    ElseIf GetAsyncKeyState(VK_S) And move_direction <> "N" Then
        check_for_keyboard_input_WASD = "S"
    ElseIf GetAsyncKeyState(VK_A) And move_direction <> "E" Then
        check_for_keyboard_input_WASD = "W"
    ElseIf GetAsyncKeyState(VK_D) And move_direction <> "W" Then '
        check_for_keyboard_input_WASD = "E"
    Else
        check_for_keyboard_input_WASD = move_direction
    End If
End Function
