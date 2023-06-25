# VBA-challenge

Sources of possible similar code snippets include:
  HW study session with Grace Palmer, Jonathan Rudamas, Adalbert Payan
  https://www.excel-easy.com/vba/examples/background-colors.html
  https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

  ChatGPT for percent formatting:
  
------
  Yes, you can convert cells to the percent format using the `NumberFormat` property with the `Cells` method in VBA. The `Cells` method allows you to refer to cells using their row and column numbers. Here's an example:

```vba
Sub ConvertToPercent()

    ' Specify the row and column numbers of the cell you want to convert
    Dim rowNumber As Long
    Dim columnNumber As Long
    rowNumber = 1
    columnNumber = 1
    
    ' Convert the cell to percent format
    Cells(rowNumber, columnNumber).NumberFormat = "0.00%"

End Sub
```

In this example, the cell in the first row and first column (`Cells(1, 1)`) is selected, and the `NumberFormat` property is set to `"0.00%"` to convert it to the percent format with two decimal places.

You can adjust the `rowNumber` and `columnNumber` variables to target the specific cell you want to convert to the percent format using the `Cells` method.

------
