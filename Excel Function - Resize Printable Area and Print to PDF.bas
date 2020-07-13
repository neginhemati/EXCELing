Attribute VB_Name = "Module8"
Sub export_as_pdf()
    
    Sheets(Array(ActiveSheet.name, "Back Cover Template")).Select
    
    Dim filename As String
    
    filename = CreateObject("WScript.Shell").specialfolders("Desktop") & "\" _
        & ActiveSheet.name
    
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim R As Range
    With ActiveSheet
        Set R = .Range("A1:G" & Last)
        .PageSetup.PrintArea = R.Address
    End With
    
    
    result = MsgBox("File will be saved as: " & filename, vbOKCancel)
    If result = vbOK Then
            
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
            filename, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
             IgnorePrintAreas:=False, OpenAfterPublish:=True
             
    End If

End Sub


