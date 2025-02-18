# Contract Communication Macro for Standard SAP Export Formatting and Consistency

## Overview
**Contract-Comm-Macro** is a simple yet powerful VBA macro that optimizes SAP export formatting for **auditable contract communication and processing**. By automating key formatting tasks, it reduces manual effort and ensures standardized output.

This macro helps:
- **Format headers and columns consistently**
- **Apply proper borders and row heights**
- **Enable filtering for better data management**
- **Reduce time spent on manual cleanup**

## Features
✅ One-click formatting via `Ctrl + K`  
✅ Automatically adjusts column widths for readability  
✅ Removes unnecessary manual adjustments  
✅ Works seamlessly with Excel exports  

## Installation & Usage
1. Open your Excel workbook.
2. Open the **Visual Basic Editor** (`Alt + F11`).
3. Import the `Contract-Comm-Macro.vba` file:
   - In the VBA Editor, go to **File > Import File...**
   - Select `Contract-Comm-Macro.vba` and click **Open**.
4. Close the editor and return to Excel.
5. Run the macro using `Ctrl + K` or via **Developer > Macros > ContractComm**.

## Code
```vba
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
    Columns("E,H:J,M").EntireColumn.AutoFit

    ' Set consistent row height for data rows
    Rows("2:500").RowHeight = 21

    ' Enable filtering for data management
    Range("A1").CurrentRegion.AutoFilter Field:=1
End Sub

Contributing

Contributions are welcome! If you’d like to improve the macro or add additional functionality:
	1.	Fork the repository
	2.	Create a feature branch (git checkout -b new-feature)
	3.	Commit your changes (git commit -m "Added new feature")
	4.	Push to the branch (git push origin new-feature)
	5.	Submit a pull request

License

This project is released under the MIT License – feel free to use, modify, and distribute it.