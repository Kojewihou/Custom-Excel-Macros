Attribute VB_Name = "SpillSafeBreakLinks"
'******************************************************************************
' Author: Elliot von Rein
' Github: https://github.com/Kojewihou
' Date: 26/08/23
' Version: 1.0
'
' Description:
' This macro corrects spills in formulas and breaks all external links in the workbook.
'
'******************************************************************************

Sub CorrectSpillsAndBreakLinks()
    Dim ws As Worksheet
    Dim Links As Variant
    Dim i As Long
    
    ' Part 1: Iterate through every worksheet and apply CorrectSpills() function
    For Each ws In ThisWorkbook.Sheets
        CorrectSpills ws
    Next ws
    
    ' Part 2: Break all external links
    Links = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(Links) Then
        For i = 1 To UBound(Links)
            ThisWorkbook.BreakLink Name:=Links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub

Sub CorrectSpills(ByVal targetWorksheet As Worksheet)
    '******************************************************************************
    ' Description:
    ' This procedure corrects spills in formulas of the given worksheet.
    '
    ' Parameters:
    ' - targetWorksheet: The worksheet to correct spills in.
    '******************************************************************************

    Dim cell As Range
    Dim formula As String
    Dim matches As String
    Dim match As Variant
    
    ' Loop through each cell in the worksheet
    For Each cell In targetWorksheet.UsedRange
        ' Check if the cell has a formula and contains "#" in the formula
        If cell.HasFormula And InStrB(1, cell.formula, "#", vbBinaryCompare) <> 0 Then
            formula = cell.formula
            matches = FindMatches(formula)
            
            ' Loop through each match and process individually
            For Each match In Split(matches, ", ")
                Dim spillCell As Range
                Dim spillRange As Range
                
                ' Get the cell specified by the match (without "#")
                Set spillCell = Range(Replace(match, "#", ""))
                
                ' Check if the spill cell has a spill
                If spillCell.HasSpill Then
                    ' Check if spill parent formula contains an external link
                    If InStr(1, spillCell.spillParent.formula, "[") > 0 Or InStr(1, spillCell.spillParent.formula, ".xl", vbTextCompare) > 0 Then
                        ' Get the spill range and replace the match in the formula
                        Set spillRange = spillCell.spillParent.SpillingToRange
                        cell.formula = Replace(cell.formula, match, spillRange.Address, , 1)
                    End If
                End If
            Next match
        End If
    Next cell
End Sub

Function FindMatches(ByVal inputFormula As String) As String
    '******************************************************************************
    ' Description:
    ' This function finds and returns matches in a given input formula using a regex pattern.
    '
    ' Parameters:
    ' - inputFormula: The formula to find matches in.
    '
    ' Returns:
    ' A comma-separated list of matched strings.
    '******************************************************************************

    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As String
    
    ' Create and configure the regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "\$?[A-Z]+\$?\d+\#"
    
    ' Find all matches in the input formula using the regular expression
    Set matches = regex.Execute(inputFormula)
    For Each match In matches
        result = result & match.value & ", "
    Next match
    
    ' Remove the trailing comma and space
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
    End If
    
    FindMatches = result
End Function

