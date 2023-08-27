Attribute VB_Name = "ApplyToAll"
'-------------------------------------------------------------
' ApplyFunctionToWorksheetCells - Apply a function to cells in all sheets
'-------------------------------------------------------------
' This macro loops through all worksheets in the workbook and applies a function
' to the cell values within each sheet's used range in chunks. It is recommended
' to adjust the chunk size based on available memory and system performance.
'
' Note: This macro affects all sheets in the workbook. Use with caution.
'
' Instructions:
' 1. Replace "AppliedFunction" with your custom function's logic.
' 2. Adjust the chunkSize value for optimal performance.
' 3. Run the macro to apply the function to all cells in all sheets.
' 4. Use Edit > Replace > Replace All to apply the function to cell formulas instead of values.
'
' Best Practices:
' - Backup your workbook before running macros that affect cell values.
' - Optimize chunkSize based on available system memory.
' - Avoid running macros on important files without proper testing.
' - Utilize Excel's Calculation and ScreenUpdating settings for efficiency.
'
' Author: Elliot von Rein
' Github: https://github.com/Kojewihou
' Date: 26/08/23
' Version: 1.0
'-------------------------------------------------------------

Sub ApplyFunctionToWorksheetCells()

    ' Disable automatic calculation and screen updating for efficiency
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim inputArray() As Variant
    Dim rowIdx As Long, colIdx As Long
    Dim chunkSize As Long
    Dim chunkStartRow As Long, chunkEndRow As Long
    
    ' Set the chunk size for data processing
    chunkSize = 100
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set the worksheet to work with
        Set rng = ws.UsedRange ' Change to the desired range
        
        ' Initialize variables for chunk processing
        chunkStartRow = rng.Row
        
        ' Loop through chunks of data within the sheet
        Do While chunkStartRow <= rng.Rows.Count
            ' Calculate the end row of the current chunk
            chunkEndRow = Application.Min(chunkStartRow + chunkSize - 1, rng.Rows.Count)
            
            ' Load cell values from the current chunk into inputArray
            inputArray = rng.Cells(chunkStartRow, rng.Column).Resize(chunkEndRow - chunkStartRow + 1, rng.Columns.Count).Value
            
            ' Apply the function to the entire inputArray in a single step
            inputArray = ApplyArrayFunction(inputArray)
            
            ' Write the output array back to the chunk's range
            rng.Cells(chunkStartRow, rng.Column).Resize(UBound(inputArray, 1), UBound(inputArray, 2)).Value = inputArray
            
            ' Move to the next chunk
            chunkStartRow = chunkEndRow + 1
        Loop
    Next ws
    
    ' Restore automatic calculation and screen updating settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

'-------------------------------------------------------------
' ApplyArrayFunction - Apply a function to a variant array
'-------------------------------------------------------------
' This function applies the specified function to each element in
' the input array and returns the modified array.
'
' Parameters:
'   inputArray (Variant) - The input array to be processed.
'
' Returns:
'   Variant - The modified array with the function applied.
'-------------------------------------------------------------
Function ApplyArrayFunction(inputArray As Variant) As Variant
    Dim rowIdx As Long, colIdx As Long

    ' Loop through the input array and apply the custom function
    For rowIdx = 1 To UBound(inputArray, 1)
        For colIdx = 1 To UBound(inputArray, 2)
            inputArray(rowIdx, colIdx) = AppliedFunction(inputArray(rowIdx, colIdx))
        Next colIdx
    Next rowIdx

    ' Return the modified array
    ApplyArrayFunction = inputArray
End Function

'-------------------------------------------------------------
' AppliedFunction - Apply your custom function's logic
'-------------------------------------------------------------
' This function represents the logic of the function you want to
' apply to each cell value. Replace the example logic with your
' actual function's implementation.
'
' Parameters:
'   inputValue (Variant) - The input value to be processed.
'
' Returns:
'   Variant - The result of the custom function's logic.
'-------------------------------------------------------------
Function AppliedFunction(inputValue As Variant) As Variant
    ' Replace this with your actual function's logic
    ' For example, cube the input value
    AppliedFunction = inputValue ^ 3
End Function

