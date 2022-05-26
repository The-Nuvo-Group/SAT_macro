VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColumnsHeaders 
   Caption         =   "Column Headers"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   OleObjectBlob   =   "ColumnsHeaders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColumnsHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function lastCl() As String

    ' Get Last Column Number with Data '
    colNumber = Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' Return the Character encoding, Column Name, of he column'
    lastCl = Split(Cells(, colNumber).address, "$")(1)

End Function

Private Function baseRow(ByVal originalRange As String) As String

    baseRow = Right(Split(originalRange, ":")(1), 2)

End Function

Private Function getRangeRowsToDelete(ByVal originalRange As String) As String
    'Array of 2 elements
    Dim rowsLimits(2) As String
    
    'Reference ->   $2:$5
    If Len(Split(originalRange, ":")(0)) = 2 Then
        rowsLimits(0) = Right(Split(originalRange, ":")(0), 1)
    End If
    If Len(Split(originalRange, ":")(0)) = 2 Then
        rowsLimits(1) = CStr(CInt(Right(Split(originalRange, ":")(1), 1)) - 1)
    End If
    
    
    'Reference ->   $12:$15
    If Len(Split(originalRange, ":")(0)) = 3 Then
        rowsLimits(0) = Right(Split(originalRange, ":")(0), 2)
    End If
    If Len(Split(originalRange, ":")(0)) = 3 Then
        rowsLimits(1) = CStr(CInt(Right(Split(originalRange, ":")(1), 2)) - 1)
    End If
    
    getRangeRowsToDelete = rowsLimits(0) & ":" & rowsLimits(1)

End Function

Private Function RowRangeToHeaderRange(ByVal originalRange As String, startRange As String, limitRange As String) As String
    Dim fst, snd As String
    fst = Left(Split(originalRange, ":")(0), 1) & startRange & "$" & Mid(Split(originalRange, ":")(0), 2, Len(Split(originalRange, ":")(0)))
    snd = Left(Split(originalRange, ":")(1), 1) & limitRange & "$" & Mid(Split(originalRange, ":")(1), 2, Len(Split(originalRange, ":")(1)))
    
    RowRangeToHeaderRange = fst & ":" & snd

End Function

Private Function getNewHdrRange(ByVal originalRange As String, startRange As String, limitRange As String) As String
    Dim toprow As String
    
    toprow = Mid(Split(originalRange, ":")(0), 2, Len(Split(originalRange, ":")(0)))
    
    getNewHdrRange = startRange & toprow & ":" & limitRange & toprow
    
End Function

Private Function LogOutputTB_run(log As String)

    'Displying Cells Specs
    LogOutputTB.WordWrap = False
    LogOutputTB.MultiLine = True
    LogOutputTB.ScrollBars = 3
    
    LogOutputTB.Value = log

End Function

Private Sub RunButton_Click()
    
    Dim hdrRange, newhdrRange As String
    hdrRange = RowRangeToHeaderRange(Split(HeadersRefEdit.Text, "!")(1), "A", lastCl())
    newhdrRange = getNewHdrRange(Split(HeadersRefEdit.Text, "!")(1), "A", lastCl())

    Dim headerData, topCell, bottomCell As Collection
    Dim txt, log, cellAddress, toDeleteRange As String
    Dim CellCounter As Integer

    Set headerData = New Collection
    Set topCell = New Collection
    Set bottomCell = New Collection
    txt = ""
    CellCounter = 1
  
    For Each Col In Range(hdrRange).Columns
        For Each Cell In Col.Cells
            If CellCounter = 1 Then
                topCell.Add Cell.address
            End If

            If CellCounter = Range(HeadersRefEdit).Rows.Count Then
                bottomCell.Add Cell.address
            End If

            txt = txt & Cell.Value & " "
            CellCounter = CellCounter + 1

            'Recording Cells specs
            log = log & "******" & vbNewLine
            log = log & "CELL: " & Cell.address & vbNewLine
            log = log & "COLOR: " & Cell.Font.Color & vbNewLine
            log = log & "SZIE: " & Cell.Font.Size & vbNewLine
            log = log & "TYPE: " & Cell.Font.Name & vbNewLine
            log = log & "BOLD: " & Cell.Font.Bold & vbNewLine
            log = log & "ALINGMENT: " & Cell.HorizontalAlignment & vbNewLine
            log = log & vbNewLine
        Next

        headerData.Add txt
        txt = ""
        CellCounter = 1
    Next

    'Display in TextBox
    LogOutputTB_run (log)
    
    'Clear Content in Range
    Range(HeadersRefEdit).ClearContents

    'Delete Rows
    Rows(getRangeRowsToDelete(Split(HeadersRefEdit.Text, "!")(1))).EntireRow.Delete
    
    'Repopulate standing Row
    Dim i As Integer
    For i = 1 To topCell.Count
        Range(topCell(i)).Value = headerData(i)
    Next
    
    'Row Specs
    With Range(newhdrRange)
    .Font.Bold = True
    .WrapText = True
    .EntireRow.AutoFit
    End With
    
    'Clear up Memory space
    Set headerData = Nothing
    Set topCell = Nothing
    Set bottomCell = Nothing
      
End Sub
