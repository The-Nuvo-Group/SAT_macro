Attribute VB_Name = "RunAnnexA1"
Sub Main()
    Dim Annex As New annexAone
    Dim Px As New PixelRatio
    
    
    Dim ws As Worksheet
    Set ws = Worksheets(2)
    ws.Activate
        
    Annex.readConfig "C:\Users\jairo\Desktop\Me\ExcelMacro\LiveRepo\Excel-Tools\Src\Config\annexa1Config.json"
    
    Annex.setupAnnexPages ws
    
    Annex.setupPrintArea ws
    
    Annex.printConfig
      
    Debug.Print "****"
    Debug.Print "****"
    Debug.Print "****"
    
    Dim charsLen, pxlen, cellLen As New Collection
    
    'For Each c In Range("A1:A145")
        'Debug.Print c.address; " - ("; Len(c.text); " - "; Len(c.Value); ") - ("; Px.getLabelPixel(c.text); " - "; Px.getLabelPixel(c.Value); ") - "; c.Width '; " - "; c.text
    'Next
    Dim targetCell As Range
    Dim cellA1 As String
    Dim idWSpace As New Collection
    
    Set targetCell = Range("A1")
    cellA1 = targetCell.text
    Set idWSpace = idxsWhiteSpaces(cellA1)
    
    Debug.Print cellA1
    Debug.Print "Numer of WhiteSpaces: "; idWSpace.Count
    
    l = Mid(cellA1, 1, idWSpace(2))
    r = Mid(cellA1, idWSpace(2), Len(cellA1))
    h = l & Chr(10) & r

    Debug.Print targetCell.address; " - ("; Len(cellA1); " - "; Len(targetCell.Value); ") - ("; Px.getLabelPixel(cellA1); " - "; Px.getLabelPixel(targetCell.Value); ") - "; targetCell.Width; " - "; cellA1
    Debug.Print Len(cellA1)
    Debug.Print Len(l)
    Debug.Print Len(r)
    Debug.Print Len(h)
    
    targetCell.Value = h
    
    Debug.Print targetCell.address; " - ("; Len(cellA1); " - "; Len(targetCell.Value); ") - ("; Px.getLabelPixel(cellA1); " - "; Px.getLabelPixel(targetCell.Value); ") - "; targetCell.Width; " - "; cellA1
        
        
        
End Sub


