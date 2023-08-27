Attribute VB_Name = "Module1"
Sub BreakLinksAndCorrectSpills()
    ' This subroutine breaks links and corrects spills in worksheets
    
    Dim ws As Worksheet
    Dim Links As Variant
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        CorrectSpillNotation ws ' Call the function to correct spills
        ' CorrectCurrentSpillWorkAround ws
    Next ws
    
    ' Breaks Links
    Links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(Links) Then
        For i = 1 To UBound(Links)
            ActiveWorkbook.BreakLink Name:=Links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub

Sub CorrectSpillNotation(ws As Worksheet)
    ' This subroutine corrects spills in the specified worksheet
    
    Dim cell As Range
    For Each cell In ws.UsedRange
        If Len(cell.Formula) > 0 Then ' Check if the cell is not empty
            ' Debug.Print cell.Formula
            ' Check if the cell matches the specified regex pattern and capture the match
            Dim matchedPattern As String
            If MatchesPattern(cell.Formula, matchedPattern) Then
            
                Dim evalCell As Range
                
                Set evalCell = Evaluate(matchedPattern)
                If evalCell.HasSpill Then
                
                    cell.Formula = Replace(cell.Formula, matchedPattern, evalCell.Address)
                    
                End If
            End If
        End If
    Next cell
End Sub



Function MatchesPattern(value As Variant, ByRef match As String) As Boolean
    ' This function checks if the value matches the specified regex pattern.
    ' If a match is found, the match is returned via the ByRef parameter.
    
    Dim regexPattern As String
    regexPattern = "\$?[A-Z]+\$?\d+\#"
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = regexPattern
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(value)
    
    If matches.Count > 0 Then
        match = matches(0).value
        MatchesPattern = True
    Else
        MatchesPattern = False
    End If
End Function

Sub CorrectCurrentSpillWorkAround(ws As Worksheet)
    ' This subroutine corrects spills in the specified worksheet
    
    Dim cell As Range
    For Each cell In ws.UsedRange
        If Len(cell.Formula) > 0 Then ' Check if the cell is not empty
            ' TODO for those with finnicky work around :P
        End If
    Next cell
End Sub
