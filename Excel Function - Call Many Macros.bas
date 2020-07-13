Attribute VB_Name = "Module7"
Sub Button_DeleteResizeRows() 'Run multiple macros by using a button:
    Call AutoAdjustRowHeight 'Macro1
    Call DeleteRowWithContents 'Macro2
End Sub
