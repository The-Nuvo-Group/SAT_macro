VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColumnsHeaders 
   Caption         =   "Column Headers"
   ClientHeight    =   1950
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

Private Sub RunButton_Click()

    Dim headerData, topCell, bottomCell As Collection
    Dim txt, cellAddress As String
    Dim counter As Integer
    
    Set headerData = New Collection
    Set topCell = New Collection
    Set bottomCell = New Collection
    txt = ""
    counter = 1
    
    For Each col In Range(HeadersRefEdit).Columns
        For Each cell In col.Cells
            If counter = 1 Then
                topCell.Add cell.address
            End If
            
            If counter = Range(HeadersRefEdit).Rows.Count Then
                bottomCell.Add cell.address
            End If
            
            txt = txt & cell.Value & " "
            counter = counter + 1
            
        Next
        Debug.Print txt
        headerData.Add txt
        txt = ""
        counter = 1
    Next
    
    'Clear Content
    Range(HeadersRefEdit).ClearContents
    
    Dim i As Long
    For i = 1 To topCell.Count
        Range(topCell(i), bottomCell(i)).Merge
        Range(topCell(i), bottomCell(i)).Value = headerData(i)
        Range(topCell(i), bottomCell(i)).WrapText = True
        
    Next
    
      
End Sub
