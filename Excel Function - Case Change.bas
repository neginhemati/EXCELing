Attribute VB_Name = "Module5"
Sub Lowercase()  'Macro name (change this based on the function of macro)
    
    Application.ScreenUpdating = False  'Temporarily shuts off screen updating for faster performance
    Application.Calculation = xlCalculationManual   'Temporarily shuts off auto calculation for faster performance
    
    For Each Cell In Selection
        If Not Cell.HasFormula Then
            Cell.Value = LCase(Cell.Value)  'Converts cell text to lowercase. For uppercase, use UCase(Cell.Value). For proper case, use  Application _ .WorksheetFunction _ .Proper(Cell.Value)
        End If
    Next Cell
    
    Application.ScreenUpdating = True  'Turns screen updating back on
    Application.Calculation = xlCalculationAutomatic 'Turns auto calculation back on

End Sub

