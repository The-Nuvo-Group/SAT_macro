Attribute VB_Name = "RunAnnexA1"
Sub Main()

    'Config files location and extension definition
    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    'Check file exists
    safeFail_configFiles configFilesLocation, cfExtension

    'Run user form
    AnnexAControlPanel.Show
             
End Sub


