Attribute VB_Name = "AnnexAdditionalTools"
Function findPropertyName(startSearch As String, endSearch As String, targetCol As String) As String
    
    Dim targetAddress As String
    Dim tmpDone As Boolean
    tmpDone = False
    'Loop Through range of cols to search'
    For Each Col In Range(startSearch, endSearch).Columns
        For Each Cell In Col.Cells
            If Cell.Value = targetCol Then
                targetAddress = Cell.address(False, False)
                tmpDone = True
                'Exit Inner loop'
                Exit For
            End If
        Next
        If tmpDone = True Then
            'Exit Outter Loop'
            Exit For
        End If
    Next
    
    'Return value'
    findPropertyName = targetAddress

End Function

Function FreezePanelTarget(startSearch As String, endSearch As String, LimitColumn As String) As String

    'Declair var Address, offSet'
    Dim address As String
    Dim FPTargetAddress As String
    
    'Get Address of Limiting Column'
    address = findPropertyName(startSearch, endSearch, LimitColumn)
    
    'Calculate Freeze Panel Address'
    FPTargetAddress = Range(address).Offset(1, 1).address(False, False)
      
    'Return Statement'
    FreezePanelTarget = FPTargetAddress

End Function

'Functions splits address, in string format, into a HashTable'
Function splitAddress(address As String) As Object
    Dim ColRow As New Dictionary
    Dim key, val
    
    'Populate HashTable with Column value'
    key = "Col": val = Right(Left(address, 2), 1)
    ColRow.Add key, val
    
    'Populate HashTable with Key value'
    key = "Row": val = Right(address, 1)
    ColRow.Add key, val
    
    'Return HashTable'
    Set splitAddress = ColRow

End Function

Function GenerateColRowTitleRange(target As String, item As String) As String
    
    Dim itemRange As String
    
    If item = "Row" Then
        itemRange = "$1:$" + target
    ElseIf item = "Col" Then
        itemRange = "$A:$" + target
    End If
    
    'Return address range'
    GenerateColRowTitleRange = itemRange

End Function

Function lastCl(ws As Worksheet) As String

    ' Get Last Column Number with Data '
    colNumber = Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' Return the Character encoding, Column Name, of he column'
    lastCl = Split(Cells(, colNumber).address, "$")(1)

End Function

Function lastRw(ws As Worksheet) As Long
    ' Get last Row with data '
    lastRw = Cells(Rows.Count, 1).End(xlUp).Row
    
End Function
