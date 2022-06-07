Attribute VB_Name = "RunAnnexA1"
Sub Main()

    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    
    'Dim Annex As New annexAone
    'Dim Px As New PixelRatio
    
    
    'Dim ws As Worksheet
    'Set ws = Worksheets(ActiveSheet.Index)
    'ws.Activate
        
    'Annex.readConfig "C:\Users\jairo\Desktop\Me\ExcelMacro\LiveRepo\Excel-Tools\Src\Config\annexa1Config.json"
    
    safeFail_configFiles configFilesLocation, cfExtension

    AnnexAControlPanel.Show
    
    'Annex.setupAnnexPages ws
    'Annex.printConfig
             
End Sub


