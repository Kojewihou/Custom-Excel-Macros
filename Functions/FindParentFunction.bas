Attribute VB_Name = "FindParentFunction"
Function FindParentCell(inputCell As Range) As Range
    Dim parentCell As Range
    Dim aboveCell As Range
    Dim leftCell As Range
    
    ' Initialize variables
    Set parentCell = inputCell
    Set aboveCell = inputCell.Offset(-1, 0)
    
    ' Move upwards while parentCell and aboveCell have spilling ranges
    Do While parentCell.HasSpill And aboveCell.HasSpill
        Set parentCell = aboveCell
        Set aboveCell = parentCell.Offset(-1, 0)
    Loop
    
    Set leftCell = parentCell.Offset(0, -1)
    
    ' Move leftwards while parentCell and leftCell have spilling ranges
    Do While parentCell.HasSpill And leftCell.HasSpill
        Set parentCell = leftCell
        Set leftCell = parentCell.Offset(0, -1)
    Loop
    
    If parentCell.HasSpill Then
        Set FindParentCell = parentCell
    Else
        Set FindParentCell = Nothing
    End If
End Function

