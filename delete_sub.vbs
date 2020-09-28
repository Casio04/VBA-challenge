
Sub delete_all()

    Application.ScreenUpdating = False
    
    For Sheet = 1 To Sheets.Count
        ActiveWorkbook.Sheets(Sheet).Columns("I:Q").Clear
        ActiveWorkbook.Sheets(Sheet).Columns("I:Q").ColumnWidth = 10.38
    Next Sheet
End Sub