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

Function FreezePanelTarget(startSearch As String, endSearch As String, LimitColumn As String) As String

    'Declair var Address, offSet'
    Dim Address As String
    Dim FPTargetAddress As String
    
    'Get Address of Limiting Column'
    Address = findPropertyName(startSearch, endSearch, LimitColumn)
    
    'Calculate Freeze Panel Address'
    FPTargetAddress = Range(Address).Offset(1, 1).Address(False, False)
      
    'Return Statement'
    FreezePanelTarget = FPTargetAddress

End Function

Sub test()
    'Activate Worksheet'
    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate
    
    Dim results As String
    results = FreezePanelTarget("A1", "H7", "Property Name")
       
    'Deactivate Freez Panes if ANY'
    ActiveWindow.FreezePanes = False
     
    Range(results).Interior.Color = RGB(255, 0, 0) 'TO BE REMOVE, FOR TESTING'
    
    'Freez  Selected Column'
    Range(results).Select
    
    'Activate Freez Panes if ANY'
    ActiveWindow.FreezePanes = True
    
End Sub
  
    


