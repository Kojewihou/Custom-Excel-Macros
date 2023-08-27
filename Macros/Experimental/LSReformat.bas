Attribute VB_Name = "Module2"
' Function to find range expressions within a formula
' Returns a collection of range expressions found in the formula
Function FindRangeExpressions(formula As String) As Collection
    ' Regular expression pattern to match range expressions like "$B$4:$O$5"
    Dim regexPattern As String
    regexPattern = "(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)"
    
    ' Create a VBScript regular expression object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configure the regular expression object
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = regexPattern
    End With
    
    ' Create a collection to store the found range expressions
    Set FindRangeExpressions = New Collection
    
    ' Execute the regular expression on the formula
    Dim matches As Object
    Set matches = regex.Execute(formula)
    
    Dim match As Object
    ' Extract and store the range expressions from the matches
    For Each match In matches
        FindRangeExpressions.Add match
    Next match
End Function

' Macro to process cells and evaluate range expressions
Sub LinkBreakSpillConversion()
    Dim ws As Worksheet
    Dim cell As Range
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Loop through each cell in the used range of the worksheet
        For Each cell In ws.UsedRange
            ' Find range expressions within the cell's formula
            Dim RangeExpressions As Collection
            Set RangeExpressions = FindRangeExpressions(cell.formula)
            
            ' If range expressions are found, call the Main function
            If RangeExpressions.Count > 0 Then
                Main cell, RangeExpressions
            End If
        Next cell
    Next ws
End Sub

Sub Main(rng As Range, RangeExpressions As Collection)
    Dim evalCell As Range
    Dim rangeFormula As Variant
    
    ' Temporarily ignore errors during loop
    'On Error Resume Next
    
    ' Loop through each range expression
    For Each rangeFormula In RangeExpressions
    
        ' Get the evaluated cell based on the range expression
        Set evalCell = rng.Worksheet.Range(rangeFormula).Cells(1)

        
        ' Check if the evaluated cell is not the same as the input cell
        If Not evalCell Is Nothing And evalCell.Address <> rng.Address Then
            
            ' Check if the evaluated cell has spill
            If evalCell.HasSpill Then
                ' Pass to reformat function
                ReformatAsLS rng, evalCell, RangeExpressions
                
                Exit For ' Exit the loop after filling once
            End If
        End If
    Next rangeFormula
    
    ' Reset error handling
    On Error GoTo 0
End Sub

Sub ReformatAsLS(cell As Range, evalCell As Range, RangeExpressions As Collection)
    Dim rangeFormula As Variant
    Dim offset As Integer
    
    ' Loop through each range expression
    For Each rangeFormula In RangeExpressions
        ' Split the formula into two parts using the rangeFormula
        Dim formulaParts As Variant
        formulaParts = Split(cell.formula, rangeFormula)
        
        ' TODO calculate offset
        offset = evalCell.Row - evalCell.spillParent.Row + 1
        
        ' It's reformatting time :0
        ' Check if formula was split into two parts
        If UBound(formulaParts) = 1 Then
            Dim partBefore As String
            partBefore = formulaParts(0)
            
            Dim partAfter As String
            partAfter = formulaParts(1)

            Dim newFormula As String
            
            If offset <= 1 Then
            
                newFormula = partBefore & "LS(" & rangeFormula & ")" & partAfter
                
            Else
                
                newFormula = partBefore & "INDEX(LS(" & rangeFormula & ")," & offset & ", 0)" & partAfter
                
            End If
            
            cell.Formula2 = newFormula
        End If
    Next rangeFormula
End Sub


