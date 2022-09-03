Attribute VB_Name = "RunAnnexA1"


'************************************************************
'Purpose:       Main subrutine that runs the macro.
'Input:         Nothing
'               Example: Main()
'Output:        Nothing
'************************************************************
Sub Main()

    'Config files location and extension definition
    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    'Check file exists
    safeFail_configFiles configFilesLocation, cfExtension

    'Run user form
    AnnexAControlPanel.Show vbModeless
             
End Sub


