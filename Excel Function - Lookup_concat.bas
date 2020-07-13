Attribute VB_Name = "Module1"
Function Lookup_concat(Search_string As String, Search_in_col As Range, Return_val_col As Range)
    
    Application.ScreenUpdating = False  'Temporarily shuts off screen updating for faster performance
    Application.Calculation = xlCalculationManual   'Temporarily shuts off auto calculation for faster performance
    
    Dim i As Long
    Dim result As String
 
    For i = 1 To Search_in_col.Count

    If Search_in_col.Cells(i, 1) = Search_string Then
    If Len(result) > 0 Then
    result = result & ", " & Return_val_col.Cells(i, 1).Value
    Else
    result = Return_val_col.Cells(i, 1).Value
    End If
    End If
    Next
 
    Lookup_concat = Trim(result)
    
    Application.ScreenUpdating = True  'Turns screen updating back on
    Application.Calculation = xlCalculationAutomatic 'Turns auto calculation back on

End Function


