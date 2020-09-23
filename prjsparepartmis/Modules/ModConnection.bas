Attribute VB_Name = "ModConnection"
Option Explicit
Public CN                       As New ADODB.Connection

Public Function Connected2DB() As Boolean
Dim isOpen As Boolean
Dim Reply  As VbMsgBoxResult

isOpen = False
On Error GoTo Err_Tracker

    Do Until isOpen = True
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Database\CMOSXP_DB.mdb" & "; Persist Security Info=False;Jet OLEDB:Database Password=qwerty123"
    CN.Open
    
    isOpen = True
    Loop
    Connected2DB = isOpen
Exit Function
Err_Tracker:
    Reply = MsgBox("Error No.:" & err.Number & vbNewLine & _
                "Description:" & err.Description, vbExclamation + vbRetryCancel, "Connection Error")
    If Reply = vbRetry Then
        Resume
    ElseIf Reply = vbCancel Then
        Connected2DB = False
    End If
End Function

