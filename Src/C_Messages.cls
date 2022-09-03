VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All variables need to be declared first
'Option Explicit

'Global Collection
Dim macroMsg As Collection


'************************************************************
'Purpose:       Constructor initializes a collection to store messages.
'Input:         Nothing
'               Example: Class_Initialize()
'Output:        Nothing
'************************************************************
Private Sub Class_Initialize()
    'Set annexA = CreateObject("Scripting.Dictionary")
    Set macroMsg = New Collection
    
End Sub


'************************************************************
'Purpose:       Stores a new messages in the collection.
'Input:         String, text to be stored
'               Example: set_message("This is a new message")
'Output:        Nothing
'************************************************************
Public Function set_message(newMsg As String)
    macroMsg.Add newMsg
End Function


'************************************************************
'Purpose:       Provides messages stored in the collection.
'Input:         Nothing
'               Example: get_messages()
'Output:        String, containing all the messages stored by order of input.
'               Example: "Message 1\n
'                         Message 2\n
'                         Message 3"
'************************************************************
Public Function get_messages() As String
    Dim msgString As String
    For Each msg In macroMsg
        msgString = msgString & msg & vbNewLine
    Next
    get_messages = msgString
    
End Function


'************************************************************
'Purpose:       Clears the queue holding the messages.
'Input:         Nothing
'               Example: clear_messages()
'Output:        Nothing
'************************************************************
Public Function clear_messages()
    Set macroMsg = New Collection
End Function
