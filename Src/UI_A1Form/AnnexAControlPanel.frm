VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnnexAControlPanel 
   Caption         =   "Annnex A-1 Control"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11220
   OleObjectBlob   =   "AnnexAControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnnexAControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global
Private configFilesName As Collection

Private Sub UserForm_Initialize()
    
    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    'Setting global var
    Set configFilesName = get_configFileNames(configFilesLocation, cfExtension)
    
    Debug.Print VarType(configFilesName(1))

    ConfigFilesMenu.Value = configFilesName(1)(0)
    ConfigFilesMenu.List = configFilesName(1)
    
End Sub

Private Sub BtnRunAnnexAMacro_Click()

    Dim Annex As New annexAone
    Dim Px As New PixelRatio
    
    
    Dim ws As Worksheet
    Set ws = Worksheets(ActiveSheet.Index)
    'ws.Activate
    
    'Get the PATH of the configFile the user selected from the dropdown menu
    Annex.readConfig CStr(configFilesName(2)(ConfigFilesMenu.ListIndex))
    
    'Annex.setupAnnexPages ws
    
    Annex.printConfig

End Sub
