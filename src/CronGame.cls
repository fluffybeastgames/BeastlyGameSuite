VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CronGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cron_body_player_1 As Collection, cron_body_player_2 As Collection
'''Public direction_player_1 As String, direction_player_2 As String

Private num_rows As Integer, num_cols As Integer
Private last_move_direction As String 'represented by "N" "E" "W" and "S"
Public game_on As Boolean
Public game_timer As Double
Public tick_duration As Double ' the closer to 0, the fast the game moves, so the denominator is a good standin for difficulty level
Public difficulty As Double ' defined as ticks per second
Const TICK_SPEED As Double = 32 ' target Ticks Per Second
Public num_players As Integer
'

Private Sub Class_Initialize()
   Debug.Print "Creating a CronGame"
   num_players = 2
   Set cron_body_player_1 = New Collection
   Set cron_body_player_2 = New Collection
   
End Sub

Public Function init(ByVal rows As Integer, ByVal cols As Integer)
   num_rows = rows
   num_cols = cols
   tick_duration = 1 / TICK_SPEED ' target seconds per game tick
    
    ' player 1
    add_cron_cell "W", 5, cols - 2, 1
    
    '''direction_player_1 = "S"
    
    ' player 2
    add_cron_cell "E", 5, 1, 2
    ''''direction_player_2 = "S"
    
        
   game_on = True
    
End Function


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

Public Sub end_game()
    If cron_body_player_1.Count > cron_body_player_2.Count Then
        MsgBox "Blue Player wins!"
    Else
        MsgBox "Pink Player wins!"
    End If
End Sub
Public Function tick(ByVal input_dir_p1 As String, ByVal input_dir_p2 As String)
    If cron_body_player_1.Count <= 0 Then
        tick = False
        Exit Function
    End If
    'Debug.Print input_dir
    
    last_row_p1 = cron_body_player_1(cron_body_player_1.Count).row
    last_col_p1 = cron_body_player_1(cron_body_player_1.Count).col
    
    'debug.Print last_row_p1
    last_row_p2 = cron_body_player_2(cron_body_player_2.Count).row
    last_col_p2 = cron_body_player_2(cron_body_player_2.Count).col
    

    'p1_added = add_cron_cell(target_row, target_col, 1)
    p1_added = add_cron_cell(input_dir_p1, last_row_p1, last_col_p1, 1)
    p2_added = add_cron_cell(input_dir_p2, last_row_p2, last_col_p2, 2)
    
    tick = p1_added And p2_added
    
End Function

Public Function add_cron_cell(input_dir, last_row, last_col, player) As Boolean ' return true if the segment was added, and false if the destination is out of bounds or already occupied
    'First determine which direction to go
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
    
    'Next, check if the move is within the board boundaries
    If target_row < 0 Or target_col < 0 Or target_row >= num_rows Or target_col >= num_cols Then
        add_cron_cell = False
        Exit Function
        
    Else
        'Then, check that the cell is not already occupied by a segment
        Dim i As Long
        For i = 1 To cron_body_player_1.Count
            'Debug.Print "Comparing against " & cron_body_player_1(i).row & ", " & cron_body_player_1(i).col
            If cron_body_player_1(i).row = target_row And cron_body_player_1(i).col = target_col Then
                add_cron_cell = False
                Exit Function
                
            End If
        Next i
        
        For i = 1 To cron_body_player_2.Count
            'Debug.Print "Comparing against " & cron_body_player_1(i).row & ", " & cron_body_player_1(i).col
            If cron_body_player_2(i).row = target_row And cron_body_player_2(i).col = target_col Then
                add_cron_cell = False
                Exit Function
                
            End If
        Next i
        ' Create a new segment and add it to the body collection
        Dim s As CronCell: Set s = New CronCell
        s.row = target_row
        s.col = target_col
        s.owner = player
        
        If player = 1 Then
            cron_body_player_1.Add s
        Else
            cron_body_player_2.Add s
        End If
        
        
        add_cron_cell = True
    
    End If
    
End Function


