Attribute VB_Name = "Module1"
Sub test()
Call replace_string("ABC", "XYZ")
End Sub
Sub replace_string(ByVal String1 As String, ByVal String2 As String)
    Dim i As Long, s As String
    With Range("A2:B1000")
        Set SearchCell = .Find(String1, LookIn:=xlValues)
        If Not SearchCell Is Nothing Then
            StartCellAddress = SearchCell.Address
            Do
                Set SearchCell = .FindNext(SearchCell)
                SearchCell.Value = String2
            Loop While Not SearchCell Is Nothing And SearchCell.Address <> StartCellAddress
        End If
    End With
End Sub
