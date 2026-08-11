VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cmdDiaContable           As ADODB.Command
Dim WithEvents RsDiaContable As ADODB.Recordset
Attribute RsDiaContable.VB_VarHelpID = -1
Dim oComando As New clsComando
    Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   
Public Function obtieneDiaContable() As Date

    Set rst1 = New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
End Function
'diaContable

Public Function obtieneFechaServidor() As Date

    Set rst1 = New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("ups_ObtieneFechaHora", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                Set rst1 = oComando.GetSP()
                obtieneFechaServidor = rst1.Fields(0)
End Function

Public Function insertarDiaContable(fdiacontable As Date) As Date

    Set rst1 = New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenInsertarDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamInput, 10, fdiacontable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                Set rst1 = oComando.GetSP()
  
End Function
'diaContable

