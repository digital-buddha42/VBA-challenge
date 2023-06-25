# VBA-challenge

Sources of possible similar code snippets include:

  Credit card example from lecture on 6/15 to mimic cell comparison for line 71

  HW study session with Grace Palmer, Jonathan Rudamas, Adalbert Payan;
    possibly their sections to calculate up to percent change
  
  https://www.excel-easy.com/vba/examples/background-colors.html;
    for formatting cell background colors in lines 92, 95.
  
  https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit;
    for autofitting column widths for cleaner look, line 161.

  ChatGPT for percent formatting:
    for formatting cells into the correct percent format "0.00%", lines 155, 159.
    Below you can find ChatGPT's reply to "can you convert a cell to a percent format in VBA?"

    
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
