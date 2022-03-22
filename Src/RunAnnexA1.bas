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
    
    'Annex.printConfig
      
    Debug.Print "****"
    Debug.Print "****"
    Debug.Print "****"
    
    Debug.Print " ------ "
'    Debug.Print Range("A1").Value = arrangeHeader(Range("A1").text)
'    Range("A1").Value = arrangeHeader(Range("A1").text)
'    Debug.Print " ------ "
'    Debug.Print arrangeHeader(Range("B1").text)
'    Range("B1").Value = arrangeHeader(Range("B1").text)
'    Debug.Print " ------ "
'    Debug.Print arrangeHeader(Range("C1").text)
'    Range("C1").Value = arrangeHeader(Range("C1").text)
'    Debug.Print " ------ "
'    Range("F1").Value = arrangeHeader(Range("F1").text)
'    Range("G1").Value = arrangeHeader(Range("G1").text)
'    Range("H1").Value = arrangeHeader(Range("H1").text)
'    Range("AA1").Value = arrangeHeader(Range("AA1").text)
'    Range("AD1").Value = arrangeHeader(Range("AD1").text)
    
    
        
End Sub


