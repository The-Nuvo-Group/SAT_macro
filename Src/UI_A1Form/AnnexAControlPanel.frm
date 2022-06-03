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


Private Sub BtnRunAnnexAMacro_Click()

    Dim Annex As New annexAone
    Dim Px As New PixelRatio
    
    
    Dim ws As Worksheet
    Set ws = Worksheets(ActiveSheet.Index)
    'ws.Activate
        
    Annex.readConfig "C:\Users\jairo\Desktop\Me\ExcelMacro\LiveRepo\Excel-Tools\Src\Config\annexa1Config.json"
    
    'AnnexAControlPanel.Show
    
    Annex.setupAnnexPages ws
    
    Annex.printConfig

End Sub

Private Sub UserForm_Initialize()
    
    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    'safeFail_configFiles configFilesLocation, cfExtension
    
    Dim configFilesName As Collection
    Set configFilesName = get_configFileNames(configFilesLocation, cfExtension)

    ConfigFilesMenu.Value = configFilesName(1)(0)
    ConfigFilesMenu.List = configFilesName(1)
    
End Sub

