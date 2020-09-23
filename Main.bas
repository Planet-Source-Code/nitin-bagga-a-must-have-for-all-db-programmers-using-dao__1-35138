Attribute VB_Name = "Module2"
Option Explicit

Global gDBType As String
Global gdbPath As String

Public Function SystemSettingsDLG() As Boolean
   Dim sForm As New SysForm
    With sForm
       SystemSettingsDLG = .Execute
    End With
    Set sForm = Nothing
End Function

Public Function InitApplication(f As Form) As Boolean
    gDBType = GetSetting(gcAppName, "DataBase", "DBType", "")
    gdbPath = GetSetting(gcAppName, "DataBase", "DBPath", "")
  
    If gDBType = "" Or gdbPath = "" Then
        If Not SystemSettingsDLG Then
            Exit Function
        Else
            gDBType = GetSetting(gcAppName, "DataBase", "DBType", "")
            gdbPath = GetSetting(gcAppName, "DataBase", "DBPath", "")
        End If
        If gDBType = "" Or gdbPath = "" Then Exit Function
    End If
    If Not ConnectToDataBase(gdbPath) Then Exit Function
    InitApplication = True
End Function

Public Function GetValue(ByVal MTable As String, ByVal MtableID As String, ByVal Id As String, ByRef Mfield As String) As String

' Pass para as table , id of table , value of id , request field of first para table
' return value as string
' Both id and requested will be string

   Dim Sql As String
   Dim ss As DAO.Recordset
   
   Sql = "SELECT " & Mfield
   Sql = Sql & " FROM " & MTable
   Sql = Sql & " Where " & MtableID & "=" & Quote(Id)
   
   Set ss = gdb.OpenRecordset(Sql, dbOpenForwardOnly)
   With ss
       If .BOF Then
           GetValue = ""
       Else
           GetValue = CkNull(ss(Mfield))
       End If
       .Close
   End With
   Set ss = Nothing
   
End Function

Public Function Quote(v As Variant, Optional c As Variant) As String
    If IsMissing(c) Then
        c = Chr(34)
    End If
    Quote = c & v & c
End Function

Public Function CkNull(v As Variant) As String
    If IsNull(v) Then
        CkNull = ""
    Else
        CkNull = CStr(v)
    End If
End Function


