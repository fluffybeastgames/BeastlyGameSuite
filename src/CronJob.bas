Attribute VB_Name = "CronJob"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Sub create_cron_sheet()
    ' First see if it exists.. if so prompt user if they want to delete it
        
    Dim boolAlreadyExists As Boolean
    boolAlreadyExists = False
    
    For Each s In ThisWorkbook.Sheets
        If s.Name = "Cron" Then
            boolAlreadyExists = True
            Exit For
        End If
    Next s
    
    If boolAlreadyExists Then
        Dim response ' alert the user and ask if they want to recreate Cron, go to the existing instance, or cancel
        response = MsgBox("There is already a sheet named 'Cron'. Click Yes to create a new instance of the game, No to navigate to Cron, or Cancel to, well, cancel.", vbYesNoCancel)
        
        If response = vbYes Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("Cron").Delete
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
        
        'Create new sheet, name it 'Cron', add game board, start button, high score, and maybe a button w/ a user form for game settings
        
        Dim wsGame As Worksheet, rngBoard As Range
        Dim rngBtnNewGame As Range, rngBtnGameSettings As Range, btnNewGame As Button, btnGameSettings As Button
            
        Set wsGame = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)) 'Create the new worksheet
        wsGame.Name = "Cron"
        
        With wsGame.Cells
            .Interior.Color = RGB(55, 55, 55)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlRight ' make it easier to read debug text
            .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        End With
        
        With wsGame.Range("B1")
            .Font.Name = "Algerian"
            .Font.Bold = True
            .Font.Size = 36
            .HorizontalAlignment = xlLeft
            .value = "CRON"
        End With
        
        Set rngBoard = wsGame.Range("C5")
        wsGame.Names.Add Name:="Game_Board_Top_Left", RefersTo:=rngBoard
        
        
        'Add New and Settings buttons and bind functionality to them
        
        Set rngBtnNewGame = wsGame.Range("B3:B3")
        Set rngBtnGameSettings = wsGame.Range("B4:B4")
                     
        Set btnNewGame = wsGame.Buttons.Add(rngBtnNewGame.Left, rngBtnNewGame.Top, rngBtnNewGame.Width, rngBtnNewGame.Height)
        Set btnGameSettings = wsGame.Buttons.Add(rngBtnGameSettings.Left, rngBtnGameSettings.Top, rngBtnGameSettings.Width, rngBtnGameSettings.Height)
        
        btnNewGame.Caption = "New Game"
        btnNewGame.Name = "BtnNewGame"
        btnNewGame.OnAction = "start_new_cron_game"
        
        btnGameSettings.Caption = "Settings"
        btnGameSettings.Name = "BtnSettings"
        btnGameSettings.OnAction = "open_cron_settings"
                     
        'Clean up
        Set wsGame = Nothing
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

Sub start_new_cron_game()

    'Gather game paramters
    Dim rows As Integer, cols As Integer, apples As Integer
    rows = 66
    cols = 99
    
    'Initiate game objects and states and start the game loop
    main_cron rows, cols
        
End Sub

Sub open_cron_settings()
    Debug.Print "open_Cron_settings"
    MsgBox "TODO"
End Sub


Sub main_cron(rows, cols)
    ''''' Setup
    Debug.Print "Initiating Cron game!"
    
    Dim engine As GameEngine: Set engine = New GameEngine
    Dim game As CronGame: Set game = New CronGame
    Dim gui As CronGUI: Set gui = New CronGUI
    
    gui.init rows, cols
    game.init rows, cols
    
    DisableArrowKeys
    
    ''' Game loop proper
    engine.run_game_loop game, gui
    
    ''''' Clean up
    EnableArrowKeys
    
    'Set wsGame = Nothing
    Set CronGUI = Nothing
    Set game = Nothing
           
End Sub


Private Sub DisableArrowKeys()
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    
    Application.OnKey "W", ""
    Application.OnKey "A", ""
    Application.OnKey "S", ""
    Application.OnKey "D", ""

    Application.OnKey "w", ""
    Application.OnKey "a", ""
    Application.OnKey "s", ""
    Application.OnKey "d", ""
    
    
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


