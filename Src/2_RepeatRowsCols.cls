VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Function finds address of target Cell (Column and Row)'
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
    
    'Return Address'
    findPropertyName = targetAddress

End Function

'Functions splits address, in string format, into a HashTable'
Function splitAddress(address As String) As Object
    Dim ColRow As Object, key, val
    Set ColRow = CreateObject("Scripting.Dictionary")
    
    'Populate HashTable with Column value'
    key = "Col": val = Left(address, 1)
    ColRow.Add key, val
    
    'Populate HashTable with Key value'
    key = "Row": val = Right(address, 1)
    ColRow.Add key, val
    
    'Return HashTable'
    Set splitAddress = ColRow

End Function

'Generate Row or Column Address Range'
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

Sub setRowCalcTitle()
    'Activate Worksheet'
    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate
    
    'Get Intersection of Rows and Columns to repeat'
    Dim results As String
    results = findPropertyName("A1", "H7", "Property Name")
    
    'Split Intersectioin into Row and Colum'
    Dim dict As Object
    Set dict = splitAddress(results)
    
    'Dummy Print'
    Debug.Print GenerateColRowTitleRange(dict.item("Row"), "Row")
    Debug.Print GenerateColRowTitleRange(dict.item("Col"), "Col")
    
    'Set Repeating Rows and Columns'
    ws.PageSetup.PrintTitleRows = GenerateColRowTitleRange(dict.item("Row"), "Row")
    ws.PageSetup.PrintTitleColumns = GenerateColRowTitleRange(dict.item("Col"), "Col")
    
    
End Sub

  


