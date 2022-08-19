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

    ConfigFilesMenu.Value = configFilesName(1)(0)
    ConfigFilesMenu.List = configFilesName(1)
    
    'Initial State of AH Flag for arranging Headers
    arrangeHeadersFlag = 0
    
End Sub

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


