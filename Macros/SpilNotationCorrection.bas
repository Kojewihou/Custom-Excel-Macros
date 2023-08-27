Attribute VB_Name = "Module2"
Sub SpillNotationCorrection()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Call the correction function for the current worksheet
        Call TheCorrectionFacility(ws)
    Next ws
End Sub


Sub TheCorrectionFacility(ws As Worksheet)
    ' Add your correction code here
    ' You can use the "ws" variable to refer to the current worksheet
    ' Perform the necessary corrections on the data within the worksheet
    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.HasFormula Then
            ' Call the Tinder function and display the matches
            Dim formula As String
            formula = cell.Formula2
            
            Dim matches As Collection
            Set matches = Tinder(formula)
            
            ' Do something with the matches
            Dim match As Variant
            For Each match In matches
                Dim targetCellAddress As String
                targetCellAddress = Split(match, ":")(0)
                
                On Error Resume Next
                Dim targetCell As Range
                Set targetCell = ws.Range(targetCellAddress)
                On Error GoTo 0
                
                If targetCell.HasSpill Then
                
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

Function Tinder(inputText As String) As Collection
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

Function Offset(targetCell As Range) As Integer
    Offset = targetCell.Row - targetCell.spillParent.Row + 1
End Function
