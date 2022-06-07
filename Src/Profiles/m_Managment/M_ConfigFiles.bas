Attribute VB_Name = "M_ConfigFiles"
'Option Explicit

Function safeFail_configFiles(ByVal configFilesLocation As String, ByVal cfExtension As String)
    Dim warningMsg As String
    
    'Test Path
    If Dir(configFilesLocation, vbDirectory) = "" Then
        warningMsg = "File Path: " & Chr(10) & Chr(10) & configFilesLocation & Chr(10) & Chr(10) & "It's not reachable, Program will end."
        MsgBox warningMsg, vbCritical
        
        'End MACRO in its ENTIRETY
        End
        
    End If
    
    'Test Config Files(AKA .json) are present
    If Dir(configFilesLocation & cfExtension) = "" Then
        warningMsg = "No Config Files (.json) present at: " & Chr(10) & Chr(10) & configFilesLocation & Chr(10) & Chr(10) & "Program will end."
        MsgBox warningMsg, vbCritical
        
        'End MACRO in its ENTIRETY
        End
        
    End If
    
End Function

Function get_configFileNames(ByVal path As String, ByVal extension As String) As Collection

    'Collection holds 2 string arrays
    Dim FilesInfo As Collection
    Set FilesInfo = New Collection
    
    'Array Definition and others
    Dim listFiles(), fullPathName() As String
    Dim strFileName, strFullPath As String
    Dim numFiles As Integer
   
   '1st reading of file names
    strFileName = Dir(path & extension)
    strFullPath = path & strFileName
    
    'Populate dynamic arrays
    Do While Len(strFileName) > 0
        ReDim Preserve listFiles(numFiles)
        ReDim Preserve fullPathName(numFiles)
        
        listFiles(numFiles) = strFileName
        fullPathName(numFiles) = strFullPath
        
        numFiles = numFiles + 1
        strFileName = Dir
        strFullPath = path & strFileName
    Loop
    
    'Populate collection with arrays
    FilesInfo.Add (listFiles)
    FilesInfo.Add (fullPathName)
        
    'Return
    Set get_configFileNames = FilesInfo
            
End Function
