Attribute VB_Name = "SpillNotationCorrection"
'******************************************************************************
' Author: Elliot von Rein
' Github: https://github.com/Kojewihou
' Date: 26/08/23
' Version: 1.0
'
' Description:
' This macro corrects spill notation in formulas on all worksheets.
'
'******************************************************************************

Sub SpillNotationCorrection()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Call the correction function for the current worksheet
        Call TheCorrectionFacility(ws)
    Next ws
End Sub

Sub TheCorrectionFacility(ByVal targetWorksheet As Worksheet)
    '******************************************************************************
    ' Description:
    ' This procedure corrects spill notation in formulas on the given worksheet.
    '
    ' Parameters:
    ' - targetWorksheet: The worksheet to correct spill notation in.
    '******************************************************************************

    Dim cell As Range
    Dim formula As String
    Dim matches As Collection
    Dim match As Variant
    
    For Each cell In targetWorksheet.UsedRange
        If cell.HasFormula Then
            formula = cell.Formula2
            Set matches = Tinder(formula)
            
            For Each match In matches
                Dim targetCellAddress As String
                targetCellAddress = Split(match, ":")(0)
                
                On Error Resume Next
                Dim targetCell As Range
                Set targetCell = targetWorksheet.Range(targetCellAddress)
                On Error GoTo 0
                
                If Not targetCell Is Nothing And targetCell.HasSpill Then
                    Dim offsetValue As Integer
                    offsetValue = Offset(targetCell)
                    
                    ' Perform your correction logic using targetCell and offsetValue
                    ' For example:
                    ' targetCell.Value = offsetValue
                    
                    Dim correctedExpression As String
                    correctedExpression = "INDEX(" & Replace(targetCell.spillParent.Address, "$", "") & "#," & offsetValue & ",0)"
                    cell.Formula2 = Replace(formula, match, correctedExpression)
                End If
            Next match
        End If
    Next cell
End Sub

Function Tinder(ByVal inputText As String) As Collection
    '******************************************************************************
    ' Description:
    ' This function finds and returns matches in a given input text using a regex pattern.
    '
    ' Parameters:
    ' - inputText: The text to find matches in.
    '
    ' Returns:
    ' A collection of matched strings.
    '******************************************************************************

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim matches As Collection
    Set matches = New Collection
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\$?[A-Z]+\$?\d+\:\$?[A-Z]+\$?\d+"
        
        Dim match As Object
        For Each match In .Execute(inputText)
            matches.Add match.value
        Next match
    End With
    
    Set Tinder = matches
End Function

Function Offset(ByVal targetCell As Range) As Integer
    '******************************************************************************
    ' Description:
    ' This function calculates the offset of a target cell within its spill range.
    '
    ' Parameters:
    ' - targetCell: The cell to calculate the offset for.
    '
    ' Returns:
    ' The offset value.
    '******************************************************************************

    Offset = targetCell.Row - targetCell.spillParent.Row + 1
End Function

