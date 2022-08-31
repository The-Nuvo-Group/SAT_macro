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
Private arrangeHeadersFlag As Integer


'************************************************************
'Purpose:       Subrutine that opens the system location to the config files
'Input:         Nothing
'               Example: ConfigFilesButton_Click("C:/SOME/PATH/IN/SYSTEM", ".json")
'Output:        Nothing
'************************************************************
Private Sub ConfigFilesButton_Click()

    Dim configFilesLocation As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    
    'Open Folder location
    Shell "C:\WINDOWS\explorer.exe """ & configFilesLocation & "", vbNormalFocus

End Sub


'************************************************************
'Purpose:       Subrutine that "displays" a multiline string in a textfield
'Input:         String, message to be display
'               Example: set_programLogRunTB2("This is a message to be display")
'Output:        Nothing
'************************************************************
Private Sub set_programLogRunTB2(msg As String)
    programLogRunTB2.WordWrap = False
    programLogRunTB2.MultiLine = True
    programLogRunTB2.ScrollBars = 3
    
    programLogRunTB2.Value = msg
End Sub


'************************************************************
'Purpose:       Subrutine that initializes the GUI form for the
'               user to interact with.
'Input:         Nothing
'               Example: UserForm_Initialize()
'Output:        Nothing
'************************************************************
Private Sub UserForm_Initialize()

    Dim configFilesLocation, cfExtension As String
    configFilesLocation = "O:\31__Nuvo Programs\Excel_ConfigFiles\"
    cfExtension = "*.json"
    
    'Setting global var
    Set configFilesName = get_configFileNames(configFilesLocation, cfExtension)

    ConfigFilesMenu.Value = configFilesName(1)(0)
    ConfigFilesMenu.List = configFilesName(1)
    
    'Initial State of AH Flag for arranging Headers
    arrangeHeadersFlag = 0
    
End Sub


'************************************************************
'Purpose:       Subrutine that runs when the 'green' button has been clicked
'               It triggers the creation of the Annex A/A-1.
'Input:         Nothing
'               Example: BtnRunAnnexAMacro_Click()
'Output:        Nothing
'************************************************************
Private Sub BtnRunAnnexAMacro_Click()

    Dim Annex As New C_annexAone
    Dim Px As New PixelRatio
    
    Dim ws As Worksheet
    Set ws = Worksheets(ActiveSheet.Index)
    
    'Get the PATH of the configFile the user selected from the dropdown menu
    Annex.readConfig CStr(configFilesName(2)(ConfigFilesMenu.ListIndex))
    
    'Check AH Flag for the current UIsession
    If arrangeHeadersFlag = 1 Then
        Annex.setArrangeHeadersFlag (1)
    Else
        Annex.setArrangeHeadersFlag (0)
    End If
    
    'Run Annex A-1 procedure
    Annex.setupAnnexPages ws
    
    'Update AG Flag for the current UI session
    arrangeHeadersFlag = Annex.getArrangeHeadersFlag
    
    'Set view to "Page Break Preview"
    ActiveWindow.View = xlPageBreakPreview
    
    'Populate TextBox
    set_programLogRunTB2 (Annex.printLogs)

End Sub


