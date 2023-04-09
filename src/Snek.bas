Attribute VB_Name = "Snek"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Sub create_snek_sheet()
    ' First see if it exists.. if so prompt user if they want to delete it
        
    Dim boolAlreadyExists As Boolean
    boolAlreadyExists = False
    
    For Each s In ThisWorkbook.Sheets
        If s.name = "Snek" Then
            boolAlreadyExists = True
            Exit For
        End If
    Next s
    
    If boolAlreadyExists Then
        Dim response ' alert the user and ask if they want to recreate snek, go to the existing instance, or cancel
        response = MsgBox("There is already a sheet named 'Snek'. Click Yes to create a new instance of the game, No to navigate to Snek, or Cancel to, well, cancel.", vbYesNoCancel)
        
        If response = vbYes Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("Snek").Delete
            Application.DisplayAlerts = True
            boolAlreadyExists = False
            
        ElseIf response = vbNo Then
            s.Activate
        
        End If
        
    End If
            
    
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
        
        Set rngBoard = wsSnek.Range("B6")
        wsSnek.Names.Add name:="Game_Board_Top_Left", RefersTo:=rngBoard
        
        
        'Add New and Settings buttons and bind functionality to them
        
        Set rngBtnNewGame = wsSnek.Range("F3:H3")
        Set rngBtnGameSettings = wsSnek.Range("J3:L3")
                     
        Set btnNewGame = wsSnek.Buttons.Add(rngBtnNewGame.Left, rngBtnNewGame.Top, rngBtnNewGame.Width, rngBtnNewGame.Height)
        Set btnGameSettings = wsSnek.Buttons.Add(rngBtnGameSettings.Left, rngBtnGameSettings.Top, rngBtnGameSettings.Width, rngBtnGameSettings.Height)
        
        btnNewGame.Caption = "New Game"
        btnNewGame.name = "BtnNewGame"
        btnNewGame.OnAction = "start_new_snek_game"
        
        btnGameSettings.Caption = "Settings"
        btnGameSettings.name = "BtnSettings"
        btnGameSettings.OnAction = "open_snek_settings"
                     
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

Sub start_new_snek_game()

    'Gather game paramters
    Dim rows As Integer, cols As Integer, apples As Integer
    rows = 10
    cols = 12
    apples = 1
    
    'Initiate game objects and states and start the game loop
    main_snek rows, cols, apples
        
End Sub

Sub open_snek_settings()
    Debug.Print "open_snek_settings"
    MsgBox "TODO"
End Sub


Sub main_snek(rows, cols, apples) 'v 2.0!
    ''''' Setup
    Debug.Print "Initiating Snek game!"
    
    Dim engine As GameEngine: Set engine = New GameEngine
    Dim game As SnekGame: Set game = New SnekGame
    Dim gui As SnekGUI: Set gui = New SnekGUI
    
    gui.init rows, cols
    game.init rows, cols, apples
    
    DisableArrowKeys
    
    ''' Game loop proper
    engine.run_game_loop game, gui
    
    ''''' Clean up
    EnableArrowKeys
    
    'Set wsSnek = Nothing
    Set SnekGUI = Nothing
    Set game = Nothing
           
End Sub


Private Sub DisableArrowKeys()
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    
'    Application.OnKey "W", ""
'    Application.OnKey "A", ""
'    Application.OnKey "S", ""
'    Application.OnKey "D", ""
'
'    Application.OnKey "w", ""
'    Application.OnKey "a", ""
'    Application.OnKey "s", ""
'    Application.OnKey "d", ""
    
    
End Sub
Sub EnableArrowKeys()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
'
    Application.OnKey "W"
    Application.OnKey "A"
    Application.OnKey "S"
    Application.OnKey "D"

    Application.OnKey "w"
    Application.OnKey "a"
    Application.OnKey "s"
    Application.OnKey "d"
End Sub

