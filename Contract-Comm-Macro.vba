Sub ContractComm()

    ' Standardizes SAP export formatting for Contract Communication Processing

    ' Keyboard Shortcut: Ctrl+K

    ' Set header formatting
    With Range("A1:BA1")
        .Interior.Color = 12611584
        .RowHeight = 42
    End With

    ' Apply consistent borders to main data range
    With Range("A1:O32")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With

    ' Optimize column widths for readability
    Columns("F").ColumnWidth = 33.22
    Columns("G:H").ColumnWidth = 8.11
    Columns("K").ColumnWidth = 10
    Columns("L").ColumnWidth = 11

    ' Auto-fit specific columns for optimal display
    Columns("E").AutoFit
    Columns("H:J").AutoFit
    Columns("M").AutoFit

    ' Set consistent row height for data rows
    Rows("2:500").RowHeight = 21

    ' Enable filtering for data management
    Range("A1").CurrentRegion.AutoFilter Field:=1

End Sub
