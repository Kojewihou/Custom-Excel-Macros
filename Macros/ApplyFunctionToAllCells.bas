Attribute VB_Name = "Module1"
' To apply function to cell formula instead of value replaye ".value" with ".formula" using edit>replace>replace all
' Adjust chunk size dependant on memory available: high memory machine can have far higher chunks


Sub ApplyFunctionToWorksheetCells()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim inputArray() As Variant
    Dim rowIdx As Long, colIdx As Long
    Dim chunkSize As Long
    Dim chunkStartRow As Long, chunkEndRow As Long

    ' Set the worksheet and range you want to work with
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Set rng = ws.UsedRange ' Change to the desired range

    ' Set the chunk size
    chunkSize = 100

    ' Initialize variables for chunk processing
    chunkStartRow = rng.Row

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
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    
End Sub

Function ApplyArrayFunction(inputArray As Variant) As Variant
    Dim rowIdx As Long, colIdx As Long

    ' Loop through the input array and apply your function
    For rowIdx = 1 To UBound(inputArray, 1)
        For colIdx = 1 To UBound(inputArray, 2)
            inputArray(rowIdx, colIdx) = AppliedFunction(inputArray(rowIdx, colIdx))
        Next colIdx
    Next rowIdx

    ApplyArrayFunction = inputArray
End Function

Function AppliedFunction(inputValue As Variant) As Variant
    ' Replace this with your actual function's logic
    AppliedFunction = inputValue * 2 ' For example, doubling the input value
End Function

