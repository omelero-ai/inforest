Attribute VB_Name = "modSeguridadInfhotel"
 Dim tLicencia As String
 Dim tFechaVence As String
 Dim Isql_Prorroga As String
 Dim RsBusca As Recordset

'vencimiento y cantidad de licencias
Public Sub obtieneVencimientoConexiones()
    If CnSeg.State = 0 Then
        CnSeg.Open
    End If
    
    tLicencia = Calcular("select isnull(tlicencia,'') as codigo from thardkey where id='" & Codificar(sRUC) & "'", CnSeg)
    If tLicencia = "0" Then
        MsgBox "Archivo de validación no encontrado", vbCritical, "Aviso"
        End
    End If
     
    tLicencia = DeCodificar(tLicencia)
    'validacion de fechas 5 ants aviso, 5 despues
    tFechaVence = Mid(tLicencia, 1, 8)
     tFechaVence = Mid(tFechaVence, 1, 2) & "/" & Mid(tFechaVence, 3, 2) & "/" & Mid(tFechaVence, 5, 4)
    sVencimientoLicencia = tFechaVence
End Sub


'=============validacionLicencias
Public Sub validacionLicenciasInfhotel()
   'para enviar correos
    Dim oCorreoElectronico As New claCorreoElectronico
    Dim mensajeProrroga As String
    

    'para validaciones Ingresos anteriores
    Dim tAcceso As String
    Dim tUltimoAcceso As String


    'para control de licencias
   
    Dim fServidor As Date
    Dim fFechaVence As Date
    Dim diferencia As Integer
    Dim nCantidadLicencias As Integer

    If CnSeg.State = 0 Then
        CnSeg.Open
    End If
    
    tLicencia = Calcular("select isnull(tlicencia,'') as codigo from thardkey where id='" & Codificar(sRUC) & "'", CnSeg)
    If tLicencia = "0" Then
        MsgBox "Archivo de validación no encontrado", vbCritical, "Aviso"
        End
    End If
     
    tLicencia = DeCodificar(tLicencia)
    'validacion de fechas 5 ants aviso, 5 despues
    tFechaVence = Mid(tLicencia, 1, 8)
    
     
    tFechaVence = Mid(tFechaVence, 1, 2) & "/" & Mid(tFechaVence, 3, 2) & "/" & Mid(tFechaVence, 5, 4)


    If tFechaVence <> "00/00/0000" Then ' diferente de ilimitado
        fFechaVence = CDate(tFechaVence)
                    'validamos fecha servidor vs ultimo ingreso
                    fServidor = Calcular("select getdate() as codigo", CnSeg)
                    tUltimoAcceso = Calcular("select isnull(tcliente,'') as codigo from thardkey where id='" & Codificar(sRUC) & "'", CnSeg)
                    If Len(tUltimoAcceso) = 51 Then
                            tUltimoAcceso = DeCodificar(tUltimoAcceso)
                            If Format(fServidor, "yyyyMMdd hh:mm:ss") < Format(tUltimoAcceso, "yyyyMMdd hh:mm:ss") Then
                                    MsgBox "Error de Ingreso. Fecha Servidor Incorrecto", vbCritical, "Aviso"
                                    End
                            End If
                    Else
                  
                            MsgBox "-Archivo de validación no encontrado", vbCritical, "Aviso"
                            End

                    End If
                    
                   
                    'calculamos diferencia de dias pendientes
                    diferencia = DateValue(fFechaVence) - DateValue(fServidor)
                    
                    'Se cambiara a 4 dias de fecha de caducidad.
        Set oCorreoElectronico = New claCorreoElectronico
        
        If diferencia < -4 Then
           
            'Cargamos los parametros de envio
            CargarParametrosCorreo
             
            If diferencia >= -7 Then
                mensajeProrroga = ", su Prorroga de pago vence en " & LTrim(str(7 + diferencia)) & " días, se envio un correo al Administrador del Local"
            Else
                mensajeProrroga = ", y su Prorroga de pago ha vencido, se envio un correo al Administrador del Local"
            End If
            
            MsgBox "Su Licencia ha caducado " & mensajeProrroga, vbCritical, "Aviso"
            
            Isql_Prorroga = "select isnull(tdetallado,'') AS tmail from TTABLA where ttabla = 'EMAIL_PRORROGA'"
            
            Set RsBusca = Lib.OpenRecordset(Isql_Prorroga, Cn)
            
            While Not RsBusca.EOF
                oCorreoElectronico.EnviarCorreoProrroga (RsBusca!tmail)
                RsBusca.MoveNext
            Wend
            
            Set RsBusca = Nothing
            
            End 'No se abre el sistema
        End If
        
        If diferencia = 0 Then
            If AvisoLicencia = 1 Then
                MsgBox "Su Letra o Cuota ha vencido. Dentro de 4 dias su Licencia Caducará", vbCritical, "Aviso"
            End If
        End If
        
        If diferencia < 0 And diferencia >= -4 Then
            If AvisoLicencia = 1 Then
                MsgBox "Su Licencia caducará en " & LTrim(str(4 + diferencia)) & " días", vbCritical, "Aviso"
            End If
        End If
        If diferencia > 0 And diferencia <= 4 Then
            If AvisoLicencia = 1 Then
                MsgBox "Su próxima Letra o Cuota vence en " & LTrim(str(diferencia)) & " días", vbCritical, "Aviso"
            End If
        End If
                    
        

'                    If diferencia < -8 Then
'                        MsgBox "Su Licencia ha caducado", vbCritical, "Aviso"
'                        End
'                    End If
'                    If diferencia = 0 Then
'                        MsgBox "Su Letra o Cuota ha vencido. Dentro de 8 dias su Licencia Caducará", vbCritical, "Aviso"
'                    End If
'                    If diferencia < 0 And diferencia >= -8 Then
'                        MsgBox "Su Licencia caducará en " & LTrim(str(8 + diferencia)) & " días", vbCritical, "Aviso"
'                    End If
'                    If diferencia > 0 And diferencia <= 8 Then
'                        MsgBox "Su próxima Letra o Cuota vence en " & LTrim(str(diferencia)) & " días", vbCritical, "Aviso"
'                    End If

    End If
    
    If sRUC <> Right(tLicencia, Len(sRUC)) Then
        MsgBox "Error 2566. Comuníquese con Infhotel Servicios Informáticos SAC", vbCritical, "Aviso"
        End
    End If

    tVersionEducativaLicencia = DeCodificar(Mid(tLicencia, 10, 1))
 
    Select Case UCase(App.EXEName)
        Case "INFOREST"
                nCantidadLicencias = Val(Mid(tLicencia, 12, 2))
        Case "ADICION"
                nCantidadLicencias = Val(Mid(tLicencia, 14, 2))
        Case "CHEFCONTROL"
                nCantidadLicencias = Val(Mid(tLicencia, 16, 2))
        Case "DESPACHADOR"
                nCantidadLicencias = Val(Mid(tLicencia, 18, 2))
        Case "ANFITRIONA"
                nCantidadLicencias = Val(Mid(tLicencia, 20, 2))
        Case "TRANFERENCIA"
                nCantidadLicencias = Val(Mid(tLicencia, 22, 2))
'        Case "ALMACEN"
'                nCantidadLicencias = Val(Mid(tLicencia, 25, 2))
'        Case "COSTOS"
'                nCantidadLicencias = Val(Mid(tLicencia, 27, 2))
'        Case "INFHOTEL"
'                nCantidadLicencias = Val(Mid(tLicencia, 30, 2))
'        Case "EVENTOS"
'                nCantidadLicencias = Val(Mid(tLicencia, 32, 2))
'        Case "PROMOCIONES"
'                nCantidadLicencias = Val(Mid(tLicencia, 34, 2))
        Case "ALMACEN"
                nCantidadLicencias = Val(Mid(tLicencia, 24, 2))
        Case "COSTOS"
                nCantidadLicencias = Val(Mid(tLicencia, 26, 2))
        Case "INFHOTEL"
                nCantidadLicencias = Val(Mid(tLicencia, 28, 2))
        Case "EVENTOS"
                nCantidadLicencias = Val(Mid(tLicencia, 30, 2))
        Case "PROMOCIONES"
                nCantidadLicencias = Val(Mid(tLicencia, 32, 2))

    End Select
    
    'si cantidad de lincencias =99 pasa y si es 00 no sale entra
    sCantidadLicenciaModulo = LTrim(str(nCantidadLicencias))
    If nCantidadLicencias = 0 Then: MsgBox "No tiene Licencias para este Módulo", vbCritical, "Aviso": End
    
    If nCantidadLicencias = 99 Then: Exit Sub
    
    Dim rsLicencias As New Recordset
    On Error GoTo ErrorSP
    Set rsLicencias = Lib.OpenRecordset("usp_Seg_cLientes '" & App.EXEName & "','" & sMDB & "'," & nCantidadLicencias & "", Cn)
    
    If Not (rsLicencias.EOF Or rsLicencias.BOF) Then
        If rsLicencias.Fields(0) = 0 Then
            MsgBox "Ha Superado el limite máximo de Licencias Permitidas para este Módulo", vbCritical, "Aviso"
            End
        End If
    End If
    fServidor = Calcular("select getdate() as codigo", CnSeg)
    tAcceso = Codificar(Format(fServidor, "yyyyMMdd hh:mm:ss"))
    
    CnSeg.Execute "update tHardKey set tCliente= '" & tAcceso & "' where id='" & Codificar(sRUC) & "'"
   
    If CnSeg.State = 1 Then
        CnSeg.Close
    End If
'
    Exit Sub
ErrorSP:
    MsgBox "Error de Activación 2566. Comuníquese con Infhotel Servicios Informáticos SAC", vbCritical, "Aviso"
    Exit Sub
End Sub
Public Sub CargarParametrosCorreo()

    Isql_Prorroga = "select isnull(tresumido,'') as Nombre ,isnull(tdetallado,'') As Valor from TTABLA where ttabla = 'EMAIL_PRORROGA_CON'"
                
    Set RsBusca = Lib.OpenRecordset(Isql_Prorroga, Cn)
    
    While Not RsBusca.EOF
    
        If RsBusca!nombre = "SMTP" Then
            Smtp_Prorroga = RsBusca!Valor
        End If
         
        If RsBusca!nombre = "Usuario" Then
            UsuarioEnvio_Prorroga = RsBusca!Valor
        End If
        
        If RsBusca!nombre = "Password" Then
            PasswordEnvio_Prorroga = RsBusca!Valor
        End If
        
        If RsBusca!nombre = "Asunto" Then
            Asunto_Prorroga = RsBusca!Valor
        End If
        
        If RsBusca!nombre = "Cuerpo" Then
            Cuerpo_Prorroga = RsBusca!Valor
        End If
        
        RsBusca.MoveNext
    Wend
    
    Set RsBusca = Nothing
    
End Sub
Private Function Codificar(strCadena As String) As String
    Dim s As String
    Dim sResultado As String
    s = UCase(strCadena)
    Text2 = ""
    For i = 1 To Len(s)
        sResultado = sResultado & "0" & Asc(Mid(s, i, 1))
    Next
    Codificar = sResultado

End Function

Private Function DeCodificar(strCadena As String) As String
    Dim s As String
    Dim sResultado As String
    s = strCadena
    Text2 = ""
    For i = 1 To Len(s) Step 3
        sResultado = sResultado & Chr(Val(Mid(s, i, 3)))
    Next
    DeCodificar = sResultado

End Function






