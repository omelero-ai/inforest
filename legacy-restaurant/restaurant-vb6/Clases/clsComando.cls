VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Cmd As ADODB.Command
Dim Rs As ADODB.Recordset
Dim ParamOK As Boolean

Public Function CreateCmdSp(cmdName As String, Cn As Connection) As Boolean
    On Error GoTo Error_CreateCmdSp
    
    CreateCmdSp = False
    Set Cmd = New ADODB.Command
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = cmdName
    Cmd.CommandTimeout = 600
    Cmd.ActiveConnection = Cn
    ParamOK = True
    CreateCmdSp = True
    Exit Function

Error_CreateCmdSp:
    MsgBox "Error Fatal: " & err.Description, vbCritical, sMensaje
End Function

Public Sub CreateParameter(paramName As String, paramType As DataTypeEnum, paramDir As ParameterDirectionEnum, paramLen As Integer, paramValue As Variant)
    On Error GoTo Error_CreateParameter
    
    Cmd.Parameters.Append Cmd.CreateParameter(paramName, paramType, paramDir, paramLen, paramValue)
    Exit Sub

Error_CreateParameter:
    ParamOK = False
End Sub

Public Function ExecSP() As Boolean
    On Error GoTo Error_ExecSP
    
    ExecSP = False
    Cmd.Execute
    ExecSP = True
    Exit Function

Error_ExecSP:
   MsgBox "Error Fatal: " & err.Description, vbCritical, sMensaje
End Function

Public Function GetSP() As ADODB.Recordset
    On Error GoTo Error_GetSp
    
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open Cmd, , adOpenStatic, adLockOptimistic
    Set GetSP = Rs
    Exit Function
    
Error_GetSp:
       MsgBox "Error Fatal: " & err.Description, vbCritical, sMensaje
End Function

Public Function GetParameterValue(paramName As String) As Variant
    GetParameterValue = Cmd.Parameters(paramName).Value
End Function

Public Sub DelSp()
    On Error Resume Next
    Set Cmd = Nothing
    Set Rs = Nothing
    ParamOK = True
End Sub

Public Function GetParamOK() As Boolean
    GetParamOK = ParamOK
End Function


