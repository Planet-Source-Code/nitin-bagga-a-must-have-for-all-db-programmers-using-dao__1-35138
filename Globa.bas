Attribute VB_Name = "Module1"
Option Explicit
Public Const gcAppName = "Friend"
Public Const gdbPath = "C:\Post\friends.mdb"
Public gWS As Workspace
Public gdb As Database
Public lCancel As Boolean
Public strValSP As String


Public Function ConnectToDataBase(dbName As String) As Boolean
    On Error GoTo MakeConnErr
    Set gWS = DBEngine.Workspaces(0)
    Set gdb = gWS.OpenDatabase(dbName, False, False)
    
    ConnectToDataBase = True
    
    On Error Resume Next
    If Err <> 0 And Err <> 3012 Then GoTo MakeConnErr
MakeConnWrapUp:
    Exit Function
MakeConnErr:
    ShowSysMsg
    ConnectToDataBase = False
    GoTo MakeConnWrapUp
End Function

Public Sub ShowSysMsg()
    Alert "Error - " & Format$(Err, "#####") & " " & Error
End Sub

Public Sub Alert(Msg)
    MsgBox Msg, vbOKOnly + vbExclamation, gcAppName
End Sub
