Attribute VB_Name = "VerticalListToArray"
Function StoreVerticalListToArray(StartCell As Range) As String()

    'Ensure that only the first cell of range is selected

    If (StartCell.Cells.Count > 1) Then
        'Select only the first cell
        StartCell = StartCell.Cells(1, 1).Address
    End If

    'End function if the first cell is empty

    If (StartCell.Value = Empty) Then
        Exit Function
    End If

    'Get the count of user e-mails listed on the distribution tab
    
    'MsgBox StartCell.Address
    
    StartCell.Range("A1").Select
    Dim ListRowCount As Integer
    ListRowCount = 0
    Do While ActiveCell.Value <> Empty
        ListRowCount = ListRowCount + 1
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    'Store e-mail list into an array
    
    StartCell.Range("A1").Select
    Dim ListArray() As String
    ReDim ListArray(ListRowCount - 1)
    For i = 0 To (ListRowCount - 1)
        ListArray(i) = ActiveCell.Value
        ActiveCell.Offset(1, 0).Activate
    Next
    
    Dim msgString As String

    'Loop to print array results for testing if necessary
    'For i = 0 To (ListRowCount - 1)
    '    msgString = ListArray(i) & vbCr
    '    MsgBox msgString
    'Next i
    
    StoreVerticalListToArray = ListArray()

End Function

