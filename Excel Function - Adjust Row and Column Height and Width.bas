Attribute VB_Name = "Module6"
Sub AutoAdjustRowHeight() 'Auto Adjusts Row Height
   Rows("11:500").AutoFit 'Autofits rows
End Sub

'Adjust Column Width - Auto
'   Sub AdjustColumnWidth() 'Adjusts Column Width
'       Columns(2).AutoFit 'Autofits 2nd column width
'       Columns("B").AutoFit 'Autofits column B
'       Columns("B").ColumnWidth = 25 'Changes column B only to 25 pixel width
'       Columns("B:E").ColumnWidth = 25 'Changes columns B:E to 25 pixel width
'   End Sub

'Adjust Row Height - Hard-coded Value
'   Sub AdjustRowHeight() 'Adjusts Row Height
'       Rows(2).AutoFit 'Autofits row 2
'       Rows(3).RowHeight = 25 'Changes row 3 only to 25 pixel height
'       Rows("3:25").RowHeight = 25 'Changes rows 3-25 to 25 pixel height
'   End Sub

