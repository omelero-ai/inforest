Attribute VB_Name = "modAuditoriaIntegral"
Public Sub registroAccesoAuditoria(strTipo As String, strUsuario As String)
On Error GoTo ErrorAuditoriaAcceso
        If CnSeg.State = 0 Then
            CnSeg.Open
        End If
        Set oComando = New clsComando
        If Not oComando.CreateCmdSp("ups_Aud_RegistraIngSalUsuario", CnSeg) Then
                Set oComando = Nothing
                Exit Sub
        End If
        oComando.CreateParameter "@tTipo", adVarChar, adParamInput, 1, strTipo
        oComando.CreateParameter "@tBaseDatos", adVarChar, adParamInput, 150, sMDB
        oComando.CreateParameter "@tModulo", adVarChar, adParamInput, 2, tModuloSeg
        oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 100, strUsuario
        oComando.CreateParameter "@nCorrelativoUsuario", adBigInt, adParamInput, 10, nCorrelativoAcceso
        oComando.CreateParameter "@nCorrelativo", adBigInt, adParamOutput, 10, nCorrelativoAcceso
            If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     GoTo ErrorAuditoriaAcceso
                     
                  Else
                     nCorrelativoAcceso = oComando.GetParameterValue("@nCorrelativo")
                  End If
            If CnSeg.State = 1 Then
                CnSeg.Close
            End If
            Exit Sub
ErrorAuditoriaAcceso:
    nCorrelativoAcceso = -1
    MsgBox "Error 2567. Comuníquese con Infhotel Servicios Informáticos SAC", vbCritical, "Aviso"
        If CnSeg.State = 1 Then
                CnSeg.Close
        End If
End Sub

 

Function RegistraMovimientoAuditoria(ByVal tModulo As String, ByVal tBaseDatos As String, ByVal tTabla As String, ByVal tDescripcionTabla As String, ByVal tAccion As String, ByVal tUsuario As String, ByVal tId_RegistroUnico As String, ByVal tId_RegistroDetalle As String, ParamArray Lista()) As Boolean
Dim rsPrimaryKeyList As New ADODB.Recordset
Dim rsColumnsList As New ADODB.Recordset
Dim rsRegistroModificado As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim ListaPK As String
Dim campo As String
Dim nCorrelativo  As String
Dim NombreCampo As String
Dim AntiguoValorCampo As String
Dim NuevoValorCampo As String
Dim nRuc As String
Dim nPais As String
'Dim CnSeg As New Connection

On Error GoTo ErrorAuditoria


        If CnSeg.State = 0 Then
            CnSeg.Open
        End If
        

nCorrelativo = Calcular("select max(nCorrelativo)+1 as Codigo from mmovimiento", CnSeg)
If nCorrelativo = 0 Then
    nCorrelativo = 1
End If

Select Case tAccion
Case "01"
    For j = 0 To ((UBound(Lista) + 1) / 3) - 1
        If j = 0 Then
           CnSeg.Execute "EXEC ups_Aud_RegistraMovimientos '" & nCorrelativo & "'," & j + 1 & ",'" & tModulo & "','" & tBaseDatos & "','" & tDescripcionTabla & "','" & tAccion & "','" & tUsuario & "','" & tId_RegistroUnico & "','" & tId_RegistroDetalle & "','" & Lista(j + 1) & "','','" & Lista(j + 2) & "'"
        Else
           CnSeg.Execute "EXEC ups_Aud_RegistraMovimientos '" & nCorrelativo & "'," & j + 1 & ",'" & tModulo & "','" & tBaseDatos & "','" & tDescripcionTabla & "','" & tAccion & "','" & tUsuario & "','" & tId_RegistroUnico & "','" & tId_RegistroDetalle & "','" & Lista((j * 2) + j + 1) & "','','" & Lista((j * 2) + j + 2) & "'"
        End If
    Next j
Case "02"
    Set rsPrimaryKeyList = Lib.OpenRecordset("SELECT COL_NAME(ic.OBJECT_ID,ic.column_id) AS ColumnName FROM sys.indexes AS i INNER JOIN sys.index_columns AS ic ON  i.OBJECT_ID = ic.OBJECT_ID AND i.index_id = ic.index_id WHERE i.is_primary_key = 1 and OBJECT_NAME(ic.OBJECT_ID)='" & tTabla & "'", Cn) ' cn es la conexion local

    Set rsColumnsList = Lib.OpenRecordset("SELECT syscolumns.name AS COL_NAME, systypes.name AS TYPE_NAME FROM syscolumns INNER JOIN sysobjects ON syscolumns.id = sysobjects.id INNER JOIN systypes ON systypes.xtype = syscolumns.xtype WHERE (sysobjects.name = '" & tTabla & "') AND (systypes.name <> 'sysname') order by colorder", Cn)
    
    'Validacion RUC ESPAÑA CSHM 28/08/2025
    nPais = Calcular("Select tcodigo as Codigo from vPaisOrigen", Cn)
    
    If rsPrimaryKeyList.RecordCount > 0 Then
        ListaPK = ""
        rsPrimaryKeyList.MoveFirst
        For i = 0 To rsPrimaryKeyList.RecordCount - 1
            For j = 0 To UBound(Lista) Step 3
                If UCase(Lista(j)) = UCase(rsPrimaryKeyList!ColumnName) Then
                    rsColumnsList.MoveFirst
                    rsColumnsList.Find "COL_NAME='" & Lista(j) & "' "
                    If Not rsColumnsList.EOF Then
                        Select Case rsColumnsList!TYPE_NAME
                            Case "varchar", "nvarchar", "text", "char", "nChar", "text", "ntext", "datetime", "smalldatetime"
                                'Validacion RUC ESPAÑA CSHM 28/08/2025
                                If nPais <> "005" Then
                                ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "='" & Lista(j + 2) & "' and "
                                Else
                                    If nPais = "005" Then
                                        nRuc = SoloNumeros(Lista(j + 2))
                                        If nRuc = "" Then
                                        ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "='" & Lista(j + 2) & "' and "
                                        Else
                                        ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "='" & nRuc & "' and "
                                        End If
                                    End If
                                End If
                            Case "bit", "bigInt", "decimal", "float", "int"
                                ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "=" & Lista(j + 2) & " and "
                        End Select
                    End If
                    Exit For
                End If
            Next j
            
            rsPrimaryKeyList.MoveNext
        Next i
        ListaPK = Left(ListaPK, Len(ListaPK) - 4)
        Set rsRegistroModificado = Lib.OpenRecordset("Select * from " & tTabla & " where " & ListaPK, Cn) ' cn es la conexion local
    Else
        'Si no hay PK no se graba el movimiento.
    End If
    
    c = 1
    For j = 0 To ((UBound(Lista) + 1) / 3) - 1 ' UBound(Lista) ' Step 3
'        If j = 115 Then
'        j = 1111
'        End If
        If j = 0 Then
           campo = Lista(j)
           NombreCampo = Lista(j + 1)
           If NombreCampo = "Id Tributario" Then
            NombreCampo = "1"
           End If
           If IsNull(rsRegistroModificado.Fields(campo).Value) Then
                AntiguoValorCampo = ""
           Else
                AntiguoValorCampo = rsRegistroModificado.Fields(campo).Value
           End If
           NuevoValorCampo = Lista(j + 2)
        Else
           campo = Lista((j * 2) + j)
           NombreCampo = Lista((j * 2) + j + 1)
           If IsNull(rsRegistroModificado.Fields(campo).Value) Then
                AntiguoValorCampo = ""
           Else
                AntiguoValorCampo = rsRegistroModificado.Fields(campo).Value
           End If
           NuevoValorCampo = Lista((j * 2) + j + 2)
        End If
        If UCase(AntiguoValorCampo) <> UCase(NuevoValorCampo) Then
            CnSeg.Execute "EXEC ups_Aud_RegistraMovimientos '" & nCorrelativo & "'," & c & ",'" & tModulo & "','" & tBaseDatos & "','" & tDescripcionTabla & "','" & tAccion & "','" & tUsuario & "','" & tId_RegistroUnico & "','" & tId_RegistroDetalle & "','" & NombreCampo & "','" & AntiguoValorCampo & "','" & NuevoValorCampo & "'"
            c = c + 1
        End If
    Next j
    
Case "03"
        If UBound(Lista) > 0 Then
            Set rsPrimaryKeyList = Lib.OpenRecordset("SELECT COL_NAME(ic.OBJECT_ID,ic.column_id) AS ColumnName FROM sys.indexes AS i INNER JOIN sys.index_columns AS ic ON  i.OBJECT_ID = ic.OBJECT_ID AND i.index_id = ic.index_id WHERE i.is_primary_key = 1 and OBJECT_NAME(ic.OBJECT_ID)='" & tTabla & "'", Cn) ' cn es la conexion local
    
            Set rsColumnsList = Lib.OpenRecordset("SELECT syscolumns.name AS COL_NAME, systypes.name AS TYPE_NAME FROM syscolumns INNER JOIN sysobjects ON syscolumns.id = sysobjects.id INNER JOIN systypes ON systypes.xtype = syscolumns.xtype WHERE (sysobjects.name = '" & tTabla & "') AND (systypes.name <> 'sysname') order by colorder", Cn)
    
            If rsPrimaryKeyList.RecordCount > 0 Then
                ListaPK = ""
                rsPrimaryKeyList.MoveFirst
                For i = 0 To rsPrimaryKeyList.RecordCount - 1
                    For j = 0 To UBound(Lista) Step 3
                        If UCase(Lista(j)) = UCase(rsPrimaryKeyList!ColumnName) Then
                            rsColumnsList.MoveFirst
                            rsColumnsList.Find "COL_NAME='" & Lista(j) & "' "
                            If Not rsColumnsList.EOF Then
                                Select Case rsColumnsList!TYPE_NAME
                                    Case "varchar", "nvarchar", "text", "char", "nChar", "text", "ntext", "datetime", "smalldatetime"
                                        ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "='" & Lista(j + 2) & "' and "
                                    Case "bit", "bigInt", "decimal", "float", "int"
                                        ListaPK = ListaPK & rsPrimaryKeyList!ColumnName & "=" & Lista(j + 2) & " and "
                                End Select
                            End If
                            Exit For
                        End If
                    Next j
                    rsPrimaryKeyList.MoveNext
                Next i
                ListaPK = Left(ListaPK, Len(ListaPK) - 4)
                Set rsRegistroModificado = Lib.OpenRecordset("Select * from " & tTabla & " where " & ListaPK, Cn)
            Else
                'Si no hay PK no se graba el movimiento.
            End If
            c = 1
                For j = 0 To ((UBound(Lista) + 1) / 3) - 1 ' UBound(Lista) ' Step 3
                    If j = 0 Then
                       campo = Lista(j)
                       NombreCampo = Lista(j + 1)
                       If IsNull(rsRegistroModificado.Fields(campo).Value) Then
                            AntiguoValorCampo = ""
                       Else
                            AntiguoValorCampo = rsRegistroModificado.Fields(campo).Value
                       End If
                    Else
                       campo = Lista((j * 2) + j)
                       NombreCampo = Lista((j * 2) + j + 1)
                        If IsNull(rsRegistroModificado.Fields(campo).Value) Then
                             AntiguoValorCampo = ""
                        Else
                             AntiguoValorCampo = rsRegistroModificado.Fields(campo).Value
                        End If
                    End If
                    CnSeg.Execute "EXEC ups_Aud_RegistraMovimientos '" & nCorrelativo & "'," & c & ",'" & tModulo & "','" & tBaseDatos & "','" & tDescripcionTabla & "','" & tAccion & "','" & tUsuario & "','" & tId_RegistroUnico & "','" & tId_RegistroDetalle & "','" & NombreCampo & "','" & AntiguoValorCampo & "',''"
                    c = c + 1
                Next j
        Else
            CnSeg.Execute "EXEC ups_Aud_RegistraMovimientos '" & nCorrelativo & "', 1,'" & tModulo & "','" & tBaseDatos & "','" & tDescripcionTabla & "','" & tAccion & "','" & tUsuario & "','" & tId_RegistroUnico & "','" & tId_RegistroDetalle & "','','',''"
        End If
End Select
RegistraMovimientoAuditoria = True
            If CnSeg.State = 1 Then
                CnSeg.Close
            End If
Exit Function
ErrorAuditoria:
            If CnSeg.State = 1 Then
                CnSeg.Close
            End If
    RegistraMovimientoAuditoria = False
    MsgBox "Error 2567. Comuníquese con Infhotel Servicios Informáticos SAC", vbCritical, "Aviso"
 
End Function


