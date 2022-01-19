VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function lastRw(ws As Worksheet) As Long
    ' Get last Row with data '
    lastRw = Cells(Rows.Count, 1).End(xlUp).Row
    
End Function

Function lastCl(ws As Worksheet) As String

    ' Get Last Column Number with Data '
    colNumber = Cells(3, Columns.Count).End(xlToLeft).Column
    
    ' Return the Character encoding, Column Name, of he column'
    lastCl = Split(Cells(, colNumber).Address, "$")(1)

End Function

Sub t()

    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate
    
    Dim Rng As Range
    Dim strRng As String
    
    
    ' -- FOR TESTING | DELETE LATER -- '
    strRng = lastCl(ws) + CStr(lastRw(ws))
    Debug.Print strRng
    
	' -- THIS RANGE IS THE ANNEX DATA -- '
    For Each Column In Range("A1", strRng).Columns
        For Each Cell In Column.Cells
			' -- FILL CELL WITH COLOR -- '
            Cell.Interior.Color = RGB(3, 202, 252)
        Next
    Next
    

End Sub


