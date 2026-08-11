VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCorreoElectronico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function EnviarCorreo(strSMTP As String, StrNombre As String, strCorreo As String, strAsunto As String, strCuerpo As String, strDe As String, strDeCorreo As String, booAdjunto As Boolean, Optional strAdjunto As String) As Boolean
    Dim mailman As ChilkatMailMan2
    Dim Email As ChilkatEmail2
    Dim contentType As String
    Dim success As Long
    Dim X As String
    
    On Error Resume Next   ' Si si se que lo odian pero aqui debe ir para evitar Crash del sistema, el exito esta validado por esta funcion
    
    Set mailman = New ChilkatMailMan2
    Set Email = New ChilkatEmail2
    
    'Registro de Licencia
    mailman.UnlockComponent ("AMeleroMAILQ_9gPfpmsvJqoF")
    mailman.SmtpHost = strSMTP
    mailman.SmtpPort = 25
    mailman.SmtpUsername = sUsuarioMail
    mailman.SmtpPassword = sClaveMail
    
    'Funcion de correccion para evitar multiples ; cuando debe ser , (coma)
    While InStr(strCorreo, ";") > 0
        strCorreo = Replace(strCorreo, ";", ",")
    Wend
    
    'email.AddTo strNombre, strNombre 'NO SE UTILIZA EL NOMBRE
    Email.AddMultipleTo (strCorreo)
    Email.subject = strAsunto
    Email.Body = strCuerpo
    
    Email.fromName = strDe
    If strDeCorreo = "" Then
        Email.FromAddress = sUsuarioMail
    Else
        Email.FromAddress = strDeCorreo
    End If
    
    If booAdjunto = True Then
        contentType = Email.AddFileAttachment(strAdjunto)
    End If

    success = mailman.SendEmail(Email)
    If success = 0 Then
        EnviarCorreo = False
        
    Else
        EnviarCorreo = True
    End If
End Function


Public Sub EnviarCorreoProrroga(strCorreo As String)
    Dim mailman As ChilkatMailMan2
    Dim Email As ChilkatEmail2
    Dim contentType As String
    Dim success As Long
    Dim X As String
               
    On Error Resume Next   ' Si si se que lo odian pero aqui debe ir para evitar Crash del sistema, el exito esta validado por esta funcion
    
    Set mailman = New ChilkatMailMan2
    Set Email = New ChilkatEmail2
    
    'Registro de Licencia
    mailman.UnlockComponent ("AMeleroMAILQ_9gPfpmsvJqoF")
    mailman.SmtpHost = Smtp_Prorroga
    mailman.SmtpPort = 25
    mailman.SmtpUsername = UsuarioEnvio_Prorroga
    mailman.SmtpPassword = PasswordEnvio_Prorroga
    
    'Funcion de correccion para evitar multiples ; cuando debe ser , (coma)
    While InStr(strCorreo, ";") > 0
        strCorreo = Replace(strCorreo, ";", ",")
    Wend
    
    'email.AddTo strNombre, strNombre 'NO SE UTILIZA EL NOMBRE
    Email.AddMultipleTo (strCorreo)
    Email.subject = Asunto_Prorroga
    Email.Body = Cuerpo_Prorroga
    
    Email.fromName = UsuarioEnvio_Prorroga
    Email.FromAddress = UsuarioEnvio_Prorroga
    
'    If booAdjunto = True Then
'        contentType = Email.AddFileAttachment(strAdjunto)
'    End If

    success = mailman.SendEmail(Email)
'    If success = 0 Then
'        EnviarCorreo = False
'    Else
'        EnviarCorreo = True
'    End If
End Sub
'
''Diego L. 03/07/2017
''Envio de Correo de Anulacion a solicitud del Hotel Tallanes.
'Public Sub EnviarCorreoAnulacion(strCorreo As String, Documento As String)
'    Dim mailman As ChilkatMailMan2
'    Dim Email As ChilkatEmail2
'    Dim contentType As String
'    Dim success As Long
'    Dim X As String
'    Dim IsqlAnulacion As String
'    Dim RsBuscaAnulacion    As ADODB.Recordset
'    Dim Usuario As String
'    Dim asunto As String
'    Dim cuerpo As String
'
'
'    On Error Resume Next   ' Si si se que lo odian pero aqui debe ir para evitar Crash del sistema, el exito esta validado por esta funcion
'
'     Set mailman = New ChilkatMailMan2
'     Set Email = New ChilkatEmail2
'
'    'Registro de Licencia
'     mailman.UnlockComponent ("AMeleroMAILQ_9gPfpmsvJqoF")
'
'        IsqlAnulacion = "select isnull(tresumido,'') as Nombre ,isnull(tdetallado,'') As Valor from TTABLA where ttabla = 'EMAIL_ANULACION'"
'
'        Set RsBuscaAnulacion = Lib.OpenRecordset(IsqlAnulacion, Cn)
'
'        While Not RsBuscaAnulacion.EOF
'
'            If RsBuscaAnulacion!nombre = "SMTP" Then
'                mailman.SmtpHost = RsBuscaAnulacion!Valor
'            End If
'
'            If RsBuscaAnulacion!nombre = "USUARIO" Then
'               Usuario = RsBuscaAnulacion!Valor
'            End If
'
'            If RsBuscaAnulacion!nombre = "PASSWORD" Then
'                mailman.SmtpPassword = RsBuscaAnulacion!Valor
'            End If
'
'            If RsBuscaAnulacion!nombre = "ASUNTO" Then
'                asunto = RsBuscaAnulacion!Valor
'            End If
'
'            If RsBuscaAnulacion!nombre = "CUERPO" Then
'                cuerpo = RsBuscaAnulacion!Valor
'            End If
'
'            RsBuscaAnulacion.MoveNext
'        Wend
'
'    Set RsBuscaAnulacion = Nothing
'
'    mailman.SmtpUsername = Usuario
'    mailman.SmtpPort = 25
'
'    'Funcion de correccion para evitar multiples ; cuando debe ser , (coma)
'    While InStr(strCorreo, ";") > 0
'        strCorreo = Replace(strCorreo, ";", ",")
'    Wend
'
'    'email.AddTo strNombre, strNombre 'NO SE UTILIZA EL NOMBRE
'    Email.AddMultipleTo (strCorreo + IIf(RsParametro!ccAnulacion = False, "", GetStr("Select ISNULL(ccEmailAnulacion,'') as Cadena FROM TPARAMETRO")))
'    Email.Subject = asunto
'    Email.Body = Replace(Replace(cuerpo, "|", vbCrLf), "@@@", ConvertirFormatoElectronico(Documento))
'
'    Email.FromName = Usuario
'    Email.FromAddress = Usuario
'
''    If booAdjunto = True Then
''        contentType = Email.AddFileAttachment(strAdjunto)
''    End If
'
'    success = mailman.SendEmail(Email)
''    If success = 0 Then
''        EnviarCorreo = False
''    Else
''        EnviarCorreo = True
''    End If
'End Sub

Function ConvertirFormatoElectronico(Documento As String) As String
Dim Serie As String
Dim Correlativo As String
Dim tdocumentoSerie As String
tdocumentoSerie = Left(Documento, 6)


    If Len(CStr(CLng(Right(tdocumentoSerie, 5)))) = 1 Then
       Serie = Left(Documento, 1) + "00" + CStr(CLng(Right(tdocumentoSerie, 5)))
    End If
    
    If Len(CStr(CLng(Right(tdocumentoSerie, 5)))) = 2 Then
      Serie = Left(Documento, 1) + "0" + CStr(CLng(Right(tdocumentoSerie, 5)))
    End If
    
    If Len(CStr(CLng(Right(tdocumentoSerie, 5)))) = 3 Then
      Serie = Left(Documento, 1) + "" + CStr(CLng(Right(tdocumentoSerie, 5)))
    End If
    
    Correlativo = CLng(Right(Documento, 6))
    
    ConvertirFormatoElectronico = Serie & "-" & Correlativo

End Function




