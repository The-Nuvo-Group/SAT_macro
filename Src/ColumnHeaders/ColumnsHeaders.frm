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

Private Function RowRangeToHeaderRange(ByVal originalRange As String, startRange As String, limitRange As String) As String
    Dim fst, snd As String
    fst = Left(Split(originalRange, ":")(0), 1) & startRange & "$" & Mid(Split(originalRange, ":")(0), 2, Len(Split(originalRange, ":")(0)))
    snd = Left(Split(originalRange, ":")(1), 1) & limitRange & "$" & Mid(Split(originalRange, ":")(1), 2, Len(Split(originalRange, ":")(1)))
    'Debug.Print fst & ":" & snd
    
    RowRangeToHeaderRange = fst & ":" & snd

End Function

Private Function LogOutputTB_run(log As String)

    'Displying Cells Specs
    LogOutputTB.WordWrap = False
    LogOutputTB.MultiLine = True
    LogOutputTB.ScrollBars = 3
    
    LogOutputTB.Value = log

End Function

Private Sub RunButton_Click()
    
    Dim hdrRange As String
    hdrRange = RowRangeToHeaderRange(Split(HeadersRefEdit.Text, "!")(1), "A", lastCl())

    Dim headerData, topCell, bottomCell As Collection
    Dim txt, log, cellAddress As String
    Dim counter As Integer

    Set headerData = New Collection
    Set topCell = New Collection
    Set bottomCell = New Collection
    txt = ""
    counter = 1

    For Each col In Range(hdrRange).Columns
        For Each cell In col.Cells
            If counter = 1 Then
                topCell.Add cell.address
            End If

            If counter = Range(HeadersRefEdit).Rows.Count Then
                bottomCell.Add cell.address
            End If

            txt = txt & cell.Value & " "
            counter = counter + 1
            
            'Recording Cells specs
            log = log & "******" & vbNewLine
            log = log & "CELL: " & cell.address & vbNewLine
            log = log & "COLOR: " & cell.Font.Color & vbNewLine
            log = log & "SZIE: " & cell.Font.Size & vbNewLine
            log = log & "TYPE: " & cell.Font.Name & vbNewLine
            log = log & "BOLD: " & cell.Font.Bold & vbNewLine
            log = log & "ALINGMENT: " & cell.HorizontalAlignment & vbNewLine
            log = log & vbNewLine
            
        Next
                
        Debug.Print txt
        headerData.Add txt
        txt = ""
        counter = 1
    Next
        
    'Display in TextBox
    LogOutputTB_run (log)


    'Clear Content in Range
    Range(HeadersRefEdit).ClearContents

    Dim i As Long
    For i = 1 To topCell.Count
        Range(topCell(i), bottomCell(i)).Merge
        Range(topCell(i), bottomCell(i)).Value = headerData(i)
        Range(topCell(i), bottomCell(i)).WrapText = True
        Range(topCell(i), bottomCell(i)).Font.Bold = True

    Next
    
    Set headerData = Nothing
    Set topCell = Nothing
    Set bottomCell = Nothing
      
End Sub
