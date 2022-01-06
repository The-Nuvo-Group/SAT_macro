VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function findPropertyName(startSearch As String, endSearch As String, targetCol As String) As String
    
    Dim targetAddress As String
    Dim tmpDone As Boolean
    tmpDone = False
    'Loop Through range of cols to search'
    For Each Col In Range(startSearch, endSearch).Columns
        For Each Cell In Col.Cells
            If Cell.Value = targetCol Then
                targetAddress = Cell.Address(False, False)
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

Sub test()
    'Activate Worksheet'
    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate

    'Declair var Address'
    Dim Address As String
    
    'Call Function'
    Address = findPropertyName("A1", "H7", "Property Name")
    
    Debug.Print Address
    
    Range(Address).Interior.Color = RGB(255, 215, 0)
    
End Sub
  
    
