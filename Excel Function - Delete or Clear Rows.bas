Attribute VB_Name = "Module4"
Sub DeleteRowWithContents() 'DELETES ALL ROWS FROM A11 DOWNWARDS WITH THE WORDs "" IN COLUMN A
    
    Application.ScreenUpdating = False  'Temporarily shuts off screen updating for faster performance
    Application.Calculation = xlCalculationManual   'Temporarily shuts off auto calculation for faster performance
    
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 11 Step -1 'Loops through last cell in range until row 11
        If (Cells(i, "A").Value) = "" Then
            Cells(i, "A").EntireRow.Delete 'Use Cells(i, "A").EntireRow.ClearContents to clear contents but not delete row
        End If
    Next i
    
    Application.ScreenUpdating = True  'Turns screen updating back on
    Application.Calculation = xlCalculationAutomatic 'Turns auto calculation back on

End Sub


