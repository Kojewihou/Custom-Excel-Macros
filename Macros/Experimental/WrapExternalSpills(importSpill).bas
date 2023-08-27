Attribute VB_Name = "Module1"
Sub WrapExternalSpills()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to the desired sheet name

    Dim highlightedParents As Collection
    Set highlightedParents = New Collection

    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.HasSpill And cell.HasFormula Then
            ' Check if the formula contains external links
            If InStr(1, cell.formula, "[") > 0 Then
                Dim spillParent As Range
                On Error Resume Next
                Set spillParent = cell.spillParent
                On Error GoTo 0
                
                If Not spillParent Is Nothing And Not Contains(highlightedParents, spillParent.Address) Then
                    Dim importFormula As String
                    importFormula = "=importSpill(" & Mid(cell.formula, 2) & ")"
                    spillParent.Formula2 = importFormula
                    highlightedParents.Add spillParent.Address
                End If
            End If
        End If
    Next cell
End Sub

Function Contains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    Contains = Not col(key) Is Nothing
    On Error GoTo 0
End Function


