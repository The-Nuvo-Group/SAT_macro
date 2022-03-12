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
    
    Debug.Print " ------ "
    Debug.Print arrangeHeader(Range("A1").text)
    Debug.Print " ------ "
    Debug.Print arrangeHeader(Range("B1").text)
    Debug.Print " ------ "
    Debug.Print arrangeHeader(Range("C1").text)
    Debug.Print " ------ "
    
    
        
End Sub


