VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnnexAControlPanel 
   Caption         =   "Annnex A-1 Control"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
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

Private Sub ConfigFilesButton_Click()

    Dim configFilesLocation As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    
    'Open Folder location
    Shell "C:\WINDOWS\explorer.exe """ & configFilesLocation & "", vbNormalFocus

End Sub

Private Sub set_programLogRunTB2(msg As String)
    programLogRunTB2.WordWrap = False
    programLogRunTB2.MultiLine = True
    programLogRunTB2.ScrollBars = 3
    
    programLogRunTB2.Value = msg
End Sub

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

    Dim Annex As New C_annexAone
    Dim Px As New PixelRatio
    
    Dim ws As Worksheet
    Set ws = Worksheets(ActiveSheet.Index)
    
    'Get the PATH of the configFile the user selected from the dropdown menu
    Annex.readConfig CStr(configFilesName(2)(ConfigFilesMenu.ListIndex))
    
    'Run Annex A-1 procedure
    Annex.setupAnnexPages ws
    
    'Populate TextBox
    set_programLogRunTB2 (Annex.printLogs)
    

End Sub
