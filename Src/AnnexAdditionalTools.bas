Attribute VB_Name = "AnnexAdditionalTools"
'************************************************************
'Purpose: 	Scales down (zoom in/out) Row hight or Columns width
'Input: 	Double, width/hight of column/row
'       	Integer, ration by which scale down is perform
'			Example:scaleDown(120.34, 40)
'Output: 	Long, new width/hight
'			Example: 48.14
'************************************************************
Function scaleDown(ByVal rcWidth As Double, ByVal ratio As Double) As Double
    Dim percent As Double
    percent = ratio / 100
    scaleDown = Format(rcWidth * percent, "0.00")
End Function


'************************************************************
'Purpose: 	Find the largest word in an array of strings
'Input: 	Array, an array of strings
'			Example: indexLargestWord(['We','are',testing','this'])
'Output: 	Integer, index of the largest word
'			Example:  2
'************************************************************
Function indexLargestWord(ByRef arrayofwords As Variant) As Integer
    Dim IDX As Integer
    Dim tmpWord As String
    
    'Find Largest Word and it index
    tmpWord = arrayofwords(0)
    For i = LBound(arrayofwords) To UBound(arrayofwords)
        If Len(arrayofwords(i)) > Len(tmpWord) Then
            tmpWord = arrayofwords(i)
            IDX = CInt(i)
        End If
    Next i
    indexLargestWord = IDX
End Function


'************************************************************
'Purpose: 	Arrange headers in a "saquare" like structure
'Input: 	Input: String, a header in string format
'			Example: arrangeHeader('We are testing this')
'Output: 	String, same header but new-lines and spaces
'			Example:  "We are\n
'					   testing\n
'					   this"
'************************************************************
Function arrangeHeader(header As String) As String

    Dim finalHeader, tmpWord As String
    Dim largest, lrgIndx, i As Integer
    Dim arrayofword() As String
    
    'If there are any "soft-enters" replace them with a space.
    header = Replace(header, Chr(10), " ")
    'Split string by spaces
    arrayofwords = Split(header)

    'If the array has more than 1 element
    If UBound(arrayofwords) > 0 Then
    
        lrgIndx = indexLargestWord(arrayofwords)
        
        'Re-arrange words to fit size
        tmpWord = ""
        For i = LBound(arrayofwords) To UBound(arrayofwords)
            'Check words len limit and loop
            If (Len(tmpWord) + Len(arrayofwords(i))) <= Len(arrayofwords(lrgIndx)) Then
                If Len(tmpWord) = 0 Then
                    tmpWord = arrayofwords(i)
                Else
                    tmpWord = tmpWord + " " + arrayofwords(i)
                End If
            Else
                'Word len limit has been reached, add to final header var & continue looping
                finalHeader = finalHeader + tmpWord + Chr(10)
                tmpWord = arrayofwords(i)
            End If
        Next
        
        'Remaining words added to the final header
        finalHeader = finalHeader + tmpWord
    
    End If
    
    arrangeHeader = finalHeader

End Function


'************************************************************
'Purpose: 	Finds a cell containing a specific string and return its address.
'Input: 	String, cell address where range should start.
'       	String, cell address where range should end.
'       	String, string to be looked in range provided.
'			Example: findPropertyName( "A1", X56", "This is a header")
'Output: 	String, address of the cell
'			Example:  "B1"
'************************************************************
Function findPropertyName(startSearch As String, endSearch As String, targetCol As String) As String
    
    Dim targetAddress As String
    Dim tmpDone As Boolean
    tmpDone = False
    'Loop Through range of cols to search'
    For Each col In Range(startSearch, endSearch).Columns
        For Each cell In col.Cells
            If cell.Value = targetCol Then
                targetAddress = cell.address(False, False)
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


'************************************************************
'Purpose: 	Locates the cell to be used at the moment of freezeing a panel. 
'Input: 	String, cell address where range should start.
'       	String, cell address where range should end.
'       	String, string to be looked in range provided.
'			Example: FreezePanelTarget( "A1", X56", "This is a header")
'Output: 	String, address of the cell
'			Example:  "C2"
'************************************************************
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


'************************************************************
'Purpose: 	Splits absolute reference address, in string format, into a HashTable
'Input: 	String, ragne address to be split
'			Example: splitAddress( "A:4")
'Output: 	String, address of the cell
'			Example:  {'Col':"A", "Row":4}
'************************************************************
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


'************************************************************
'Purpose: 	Generates a Range address for the data headers.
'Input: 	String, target or limit of the address
'			String, type of address. Either a column or a row address
'			Example: GenerateColRowTitleRange("4","Col")
'Output: 	String, address of the range
'			Example:  "$A:$4"
'************************************************************
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


'************************************************************
'Purpose: 	Find the last column with data in the sheet
'Input: 	Worksheet, active worksheet
'			Example: lastCl(active-worksheet)
'Output: 	String, column letter
'			Example:  "AA"
'************************************************************
Function lastCl(ws As Worksheet) As String

    ' Get Last Column Number with Data '
    colNumber = Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' Return the Character encoding, Column Name, of he column'
    lastCl = Split(Cells(, colNumber).address, "$")(1)

End Function


'************************************************************
'Purpose: 	Find the last row with data in the sheet
'Input: 	Worksheet, active worksheet
'			Example: lastRw(active-worksheet)
'Output: 	String, row number
'			Example:  "222"
'************************************************************
Function lastRw(ws As Worksheet) As Long
    ' Get last Row with data '
    lastRw = Cells(Rows.Count, 1).End(xlUp).Row
    
End Function


'************************************************************
'Purpose: 	Finds the indexes of white spaces in a string
'Input: 	String, text to be parsed
'			Example: idxsWhiteSpaces('We are testing this')
'Output: 	String, row number
'			Example:  [2,6,13]
'************************************************************
'Function returns the idx location of a White Space in a String
Function idxsWhiteSpaces(text As String) As Collection
    Dim idxs As New Collection
    Dim IDX As Integer
    IDX = 0
    Do While True
        IDX = InStr(IDX + 1, text, " ", vbBinaryCompare)
        If IDX = 0 Then
            Exit Do
        End If
        idxs.Add IDX
    Loop
    
    'Return
    Set idxsWhiteSpaces = idxs
End Function


