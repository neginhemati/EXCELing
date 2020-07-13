Attribute VB_Name = "Module2"
Public Sub CopySelectedSheets()
    ActiveWindow.SelectedSheets.Copy Before:=Worksheets("Back Cover Template") 'Copies active sheet to before hard-coded workbook
    
    On Error Resume Next 'Skips error to rename sheet
    ActiveSheet.Name = "New PRT"
End Sub
    

'Source: https://www.ablebits.com/office-addins-blog/2018/12/05/duplicate-sheet-excel-vba/
'More options additional to the below are available at the source

'Copy sheet to new workbook:
'   Public Sub CopySheetToNewWorkbook()
'       ActiveSheet.Copy
'   End Sub


'Copy multiple sheets to new workbook:
'   Public Sub CopySelectedSheets()
'       ActiveWindow.SelectedSheets.Copy
'   End Sub


'Copy sheet to the beginning of another workbook:
'   Public Sub CopySheetToBeginningAnotherWorkbook()
'       ActiveSheet.Copy Before:=Workbooks("Book1.xlsx").Sheets(1)
'   End Sub


'Copy sheet to the end of another workbook:
'   Public Sub CopySheetToEndAnotherWorkbook()
        'ActiveSheet.Copy After:=Workbooks("Book1.xlsx").Sheets(Workbooks("Book1.xlsx").Worksheets.Count)
'   End Sub


'Copy sheet to a selected workbook:
'See code at to https://www.ablebits.com/office-addins-blog/2018/12/05/duplicate-sheet-excel-vba/
    
    
'To allow the user to specify the name for the copied sheet:
'   Public Sub CopySheetAndRename()
'       Dim newName As String
 
'       On Error Resume Next
'       newName = InputBox("Enter the name for the copied worksheet")
 
'       If newName <> "" Then
'           ActiveSheet.Copy After:=Worksheets(Sheets.Count)
'           On Error Resume Next
'           ActiveSheet.Name = newName
'       End If
'   End Sub


'To copy sheet and rename based on cell value:
'   Public Sub CopySheetAndRenameByCell2()
'       Dim wks As Worksheet
'       Set wks = ActiveSheet
'       ActiveSheet.Copy After:=Worksheets(Sheets.Count)
'       If wks.Range("A1").Value <> "" Then
'           On Error Resume Next
'           ActiveSheet.Name = wks.Range("A1").Value
'       End If
'       wks.Activate
'   End Sub
