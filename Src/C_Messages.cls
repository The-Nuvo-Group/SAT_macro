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

Private Sub Class_Initialize()
    'Set annexA = CreateObject("Scripting.Dictionary")
    Set macroMsg = New Collection
    
End Sub


Public Function set_message(newMsg As String)
    macroMsg.Add newMsg
End Function

Public Function get_messages() As String
    Dim msgString As String
    For Each msg In macroMsg
        msgString = msgString & msg & vbNewLine
    Next
    get_messages = msgString
    
End Function

Public Function clear_messages()
    Set macroMsg = New Collection
End Function
