VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnekGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public snek_body As Collection
Public apples As New Collection
Private num_rows As Integer, num_cols As Integer
Private last_move_direction As String 'represented by "N" "E" "W" and "S"
Public game_on As Boolean
Public game_timer As Double
Public tick_duration As Double ' the closer to 0, the fast the game moves, so the denominator is a good standin for difficulty level
Public difficulty As Double ' defined as ticks per second
Const DIFFUCLTY_STARTING As Double = 2.75
Const DIFFUCLTY_MAX As Double = 10
'Const num_players As Integer = 1
Public num_players As Integer

'
Private Sub Class_Initialize()
   Debug.Print "Creating a SnekGame"
   
   Set snek_body = New Collection
   Set apples = New Collection
   num_players = 1

End Sub

Public Function init(ByVal rows As Integer, ByVal cols As Integer, ByVal num_apples As Integer)
   num_rows = rows
   num_cols = cols
   difficulty = DIFFUCLTY_STARTING
   tick_duration = 1 / difficulty ' target seconds per game tick
    
   add_snek_segment 1, 1
   
   For i = 1 To num_apples
       add_apple
   Next i
      
    
   game_on = True
    
End Function

Public Sub end_game()
    MsgBox "Thanks for playing! Your score is " & snek_body.Count - 1 & ".", vbOKOnly, "Game Over"
'    response = MsgBox("Thanks for playing! Your score is " & snek_body.Count - 1 & ".", vbRetryCancel, "Game Over")
'    If response = vbRetry Then
'        MsgBox "TO DO - new game :)"
'
'    End If
    
End Sub
Property Get cols() As Integer
    cols = num_cols
End Property

Property Let cols(ByVal cols As Integer)
    num_cols = cols
End Property

Property Get rows() As Integer
    rows = num_rows
End Property

Property Let rows(ByVal rows As Integer)
    num_rows = rows
End Property


Public Function tick(ByVal input_dir As String)
    If snek_body.Count <= 0 Then
        tick = False
        Exit Function
    End If
        
    last_row = snek_body(snek_body.Count).row
    last_col = snek_body(snek_body.Count).col
        
    If input_dir = "N" Then
        target_row = last_row - 1
        target_col = last_col
        
    ElseIf input_dir = "S" Then
        target_row = last_row + 1
        target_col = last_col
        
    ElseIf input_dir = "W" Then
        target_row = last_row
        target_col = last_col - 1
        
    ElseIf input_dir = "E" Then
        target_row = last_row
        target_col = last_col + 1
    
    Else
        Debug.Print "ERROR invalid input direction"
    End If
    
    'Debug.Print "Moving " & input_dir & " from " & last_row & ", " & last_col & " to " & target_row & ", " & target_col
    If valid_move(target_row, target_col) Then
        add_snek_segment target_row, target_col
        
        If snek_body.Count Mod 4 = 0 Then 'increase the difficulty
            increase_difficulty
        End If
        
        tick = True
    Else
        'End game
        Debug.Print "INVALID"
        tick = False
        
    End If
    


End Function

Private Sub increase_difficulty()
    If difficulty < DIFFUCLTY_MAX Then
        difficulty = difficulty + 0.5
        tick_duration = 1 / difficulty
    End If
End Sub

Public Function add_snek_segment(row, col) As Boolean ' return true if the segment was added, and false if the destination is out of bounds or already occupied
    'Debug.Print "Add snek seg to " & row & ", " & col
    'First, check if the move is within the board boundaries
    If row < 0 Or col < 0 Or row >= num_rows Or col >= num_cols Then
        add_snek_segment = False
        Exit Function
        
    Else
        'Second, check that the cell is not already occupied by a segment
        Dim i As Long
        For i = 1 To snek_body.Count
            'Debug.Print "Comparing against " & snek_body(i).row & ", " & snek_body(i).col
            If snek_body(i).row = row And snek_body(i).col = col Then
                add_snek_segment = False
                Exit Function
                
            End If
        Next i
        
        ' Create a new segment and add it to the body collection
        Dim s As SnekSegment: Set s = New SnekSegment
        s.row = row
        s.col = col
        snek_body.Add s
        
        'Check if there was an apple in the new cell. If yes, remove the apple. If no, remove the oldest snake segment
        Dim apple_present As Boolean: apple_present = False
        For i = 1 To apples.Count
            If apples(i).row = row And apples(i).col = col Then
                apple_present = True
                apples.Remove i
                Exit For
            End If
        Next i
        
        If apple_present Then
            add_apple
        ElseIf snek_body.Count > 1 Then
            snek_body.Remove 1

        End If
        
        add_snek_segment = True
    
    End If
    
End Function


Function add_apple() As Boolean ' Returns true unless the board is completely full
    If apples.Count + snek_body.Count >= num_rows * num_cols Then ' Make sure there's room before we start
        add_apple = False
        Exit Function
    End If
    
    
    Dim a As SnekApple
    Dim row As Integer, col As Integer
    
    Dim empty_spot_found As Boolean, no_segments_in_cell As Boolean, no_apples_in_cell As Boolean '
    
    empty_spot_found = False
    
    Do While Not empty_spot_found
        row = Rnd * (num_rows - 1) ' a more sophisticated approach to remove the guess work - compile an array of available cells and pick a random element of the array
        col = Rnd * (num_cols - 1)
        
        no_segments_in_cell = True ' until proven otherwise
        no_apples_in_cell = True ' until proven otherwise
                
        For i = 1 To snek_body.Count
            
            If snek_body(i).row = row And snek_body(i).col = col Then
                no_segments_in_cell = False
                Exit For
            End If
        Next i
        
        For i = 1 To apples.Count
            If apples(i).row = row And apples(i).col = col Then
                no_apples_in_cell = False
                Exit For
            End If
        Next i
        
        
        If no_segments_in_cell And no_apples_in_cell Then
            empty_spot_found = True
            
            Set a = New SnekApple
            a.row = row
            a.col = col
            apples.Add a
        End If
    Loop
    
    add_apple = True
    
End Function


Private Function valid_move(target_row, target_col)
    valid_move = True ' until proven otherwise
    
    If target_row < 0 Or target_row >= num_rows Or target_col < 0 Or target_col >= num_cols Then 'check out of bounds
        valid_move = False
    Else
        For Each seg In snek_body 'check for collisions
            If target_row = seg.row And target_col = seg.col Then
                valid_move = False
                Exit Function
            End If
        Next seg
    
    End If

End Function

