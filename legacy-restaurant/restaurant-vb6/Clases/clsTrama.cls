VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim txtTrama As String
' Datos de los ini
Dim cantHeader As Integer
Dim Separador, PrefIni, PrefFin As String

Private Sub DatosTrama(TipoIni As String)
On Error GoTo fin:
    cantHeader = CInt(Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", "Datos", "ITEMHEADER", ".")))
    Separador = (Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", "Datos", "SEPARADOR", ".")))
    If TipoIni <> "DOC" Then
        PrefIni = (Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", "Datos", "PREFINI", "")))
        PrefFin = (Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", "Datos", "PREFFIN", "")))
    End If
    Exit Sub
fin:
    Call Log_Inforest("Paperlees", "DatosTrama", TipoIni, "Caja: " + sCaja, "", error, "", "FALLA PROCESO", sUsuario)
End Sub
Private Function ArmarDato(TipoIni As String, HeaderTail As String, Item As String, dato As String) As String
On Error GoTo fin
    Dim Large, posicion, tamaño As Integer
    Dim DataType As String
    Dim Valor As String
    ArmarDato = dato
    
    ArmarDato = Replace(ArmarDato, "|", "")

    Valor = (Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", HeaderTail, "ITEM" & Item, ".")))
    
    posicion = InStr(1, Valor, "|")
    tamaño = Len(Trim(Valor))
    
    DataType = CStr(Mid(Valor, 1, posicion - 1))
    Large = CInt(Mid(Valor, posicion + 1, tamaño))
    
    If DataType = "a" Then ' caracter alfabetico
        If ValAlfabetico(dato) = "" Then
            GoTo fin:
        End If
    ElseIf DataType = "n" Then ' caracter numerico
        If Not IsNumeric(dato) Then
            GoTo fin:
        End If
    ElseIf DataType = "an" Then ' caracter alfanumerico
        
    Else ' no se encontro ningun caracter
        GoTo fin:
    End If
    
    If ArmarDato <> "" Then
        If Large > 0 Then
            If Len(ArmarDato) > Large Then
                ArmarDato = Left(ArmarDato, Large)
            ElseIf Len(ArmarDato) < Large Then
'                If InStr(1, ArmarDato, "&") > 0 Or InStr(1, ArmarDato, "<") > 0 Then
'                    ArmarDato = "<![CDATA[" & ArmarDato & "]]>"
'                Else
                    ArmarDato = String(Large - Len(ArmarDato), " ") & ArmarDato
                'End If
                
            End If
        End If
    End If

    Exit Function
fin:
    ArmarDato = ""
    Call Log_Inforest("Paperlees", "ArmarDato", TipoIni, "Caja: " + sCaja, HeaderTail, error, Item, dato, sUsuario)
End Function
Private Function ValAlfabetico(Valor As String) As String
On Error GoTo fin
    ValAlfabetico = Valor
    If IsNumeric(Valor) Then
        ValAlfabetico = ""
        Exit Function
    End If
    If InStr(1, Valor, "0") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "1") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "2") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "3") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "4") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "5") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "6") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "7") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "8") > 0 Then ValAlfabetico = "": Exit Function
    If InStr(1, Valor, "9") > 0 Then ValAlfabetico = "": Exit Function
Exit Function
fin:
ValAlfabetico = ""
Call Log_Inforest("Paperlees", "ValAlfabetico", "", "Caja: " + sCaja, Valor, error, "", "FALLA PROCESO", sUsuario)
End Function

Public Function EncriptaTrama(Datos As Recordset, NumTail As String) As String
Dim i As Integer
Dim texto As String
On Error GoTo fin
        Call CrearCarpetas(App.Path + "\ConfiguracionTrama\Trama\")
        Call DatosTrama("SEND")
        If Datos.RecordCount > 0 Then
            Datos.MoveFirst
            ' cabecera de la Mensajeria------------------------------------------------------------------------------------------------------
            For i = 0 To cantHeader - 1
                If Not (Datos.BOF Or Datos.EOF) Then
                    If Datos.Fields.Count - 1 > i Then
                        texto = texto & ArmarDato("SEND", "Header", CStr(i), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                    End If
                Else
                    EncriptaTrama = ""
                    Exit Function
                End If
            Next i
            ' detalle de la mensajeria------------------------------------------------------------------------------------------------------------
            Dim ENTRAR As Integer
            ENTRAR = 0
            If cantHeader <= Datos.Fields.Count - 1 Then
                Dim ii As Integer
                ii = 0
                Do Until Datos Is Nothing
                    If ENTRAR = 0 Then ' Detalle del primer recorset enviada----------------------------------------------------------------
                        For i = cantHeader To Datos.Fields.Count - 1
                            If Not (Datos.BOF Or Datos.EOF) Then
                                If Datos.Fields.Count - 1 > i Then
                                    texto = texto & ArmarDato("SEND", "Tail" & NumTail, CStr(ii), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                                ElseIf Datos.Fields.Count - 1 = i Then
                                    texto = texto & ArmarDato("SEND", "Tail" & NumTail, CStr(ii), Datos.Fields(i)) '& IIf(Separador = "tab", vbTab, Separador)
                                End If
                            Else
                                EncriptaTrama = ""
                                Exit Function
                            End If
                            ii = ii + 1
                        Next i
                        ENTRAR = 1
                    Else ' detalle de los siguienteas recoraset enviados como informacion----------------------------------------------------
                        Call DatosTrama("DOC")
                        Do While Not Datos.EOF
                            If Datos.Fields.Count > 0 Then
                                For i = 0 To Datos.Fields.Count - 1
                                    If Datos.Fields.Count - 1 > i Then
                                        If Datos.Fields(i) = vbNewLine Then
                                            texto = Mid(texto, 1, Len(Trim(texto)) - 1)
                                            texto = texto & vbNewLine
                                        Else
                                            texto = texto & ArmarDato("DOC", "Tail" & CStr(ENTRAR), CStr(i), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                                        End If
                                    ElseIf Datos.Fields.Count - 1 = i Then
                                        If Datos.Fields(i) = vbNewLine Then
                                            texto = Mid(texto, 1, Len(Trim(texto)) - 1)
                                            texto = texto & vbNewLine
                                        Else
                                            texto = texto & ArmarDato("DOC", "Tail" & CStr(ENTRAR), CStr(i), Datos.Fields(i))
                                        End If
                                    End If
                                Next i
                            End If
                            Datos.MoveNext
                        Loop
                        ENTRAR = ENTRAR + 1
                    End If
                    Set Datos = Datos.NextRecordset()
                Loop
            End If
            EncriptaTrama = PrefIni + texto + PrefFin
        End If
   EncriptaTrama = Trim(EncriptaTrama)
Exit Function
fin:
    EncriptaTrama = ""
    Call Log_Inforest("Paperlees", "EncriptaTrama", " Revsar SP TCPIP", "Caja: " + sCaja, NumTail, error, "", "FALLA PROCESO", sUsuario)
End Function
Public Function VolcarTrama(Datos As Recordset, NumTail As String) As String
Dim i As Integer
Dim texto As String
On Error GoTo fin

        Call DatosTrama("VOLCADO")
        If Datos.RecordCount > 0 Then
            Datos.MoveFirst
            ' cabecera de la Mensajeria------------------------------------------------------------------------------------------------------
            For i = 0 To cantHeader - 1
                If Not (Datos.BOF Or Datos.EOF) Then
                    If Datos.Fields.Count - 1 > i Then
                        texto = texto & ArmarDato("VOLCADO", "Header", CStr(i), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                    End If
                Else
                    VolcarTrama = ""
                    Exit Function
                End If
            Next i
            ' detalle de la mensajeria------------------------------------------------------------------------------------------------------------
            Dim ENTRAR As Integer
            ENTRAR = 0
            If cantHeader <= Datos.Fields.Count - 1 Then
                Dim ii As Integer
                ii = 0
                Do Until Datos Is Nothing
                    If ENTRAR = 0 Then ' Detalle del primer recorset enviada----------------------------------------------------------------
                        Do While Not Datos.EOF
                            ii = 0
                            For i = cantHeader To Datos.Fields.Count - 1
                                If Not (Datos.BOF Or Datos.EOF) Then
                                    If Datos.Fields.Count - 1 > i Then
                                        If Datos.Fields(i) = vbNewLine Then
                                            texto = Mid(texto, 1, Len(Trim(texto)) - 1)
                                            texto = texto & vbNewLine
                                        Else
                                            texto = texto & ArmarDato("VOLCADO", "Tail" & NumTail, CStr(ii), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                                        End If
                                        'texto = texto & ArmarDato("VOLCADO", "Tail" & NumTail, CStr(ii), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
                                    ElseIf Datos.Fields.Count - 1 = i Then
                                        If Datos.Fields(i) = vbNewLine Then
                                            texto = Mid(texto, 1, Len(Trim(texto)) - 1)
                                            texto = texto & vbNewLine
                                        Else
                                            texto = texto & ArmarDato("VOLCADO", "Tail" & NumTail, CStr(ii), Datos.Fields(i))
                                        End If
                                        'texto = texto & ArmarDato("VOLCADO", "Tail" & NumTail, CStr(ii), Datos.Fields(i)) '& IIf(Separador = "tab", vbTab, Separador)
                                    End If
                                Else
                                    VolcarTrama = ""
                                    Exit Function
                                End If
                                ii = ii + 1
                            Next i
                            Datos.MoveNext
                        Loop
'                        ENTRAR = 1
'                    Else ' detalle de los siguienteas recoraset enviados como informacion----------------------------------------------------
'                        Call DatosTrama("DOC")
'                        Do While Not Datos.EOF
'                            If Datos.Fields.Count > 0 Then
'                                For i = 0 To Datos.Fields.Count - 1
'                                    If Datos.Fields.Count - 1 > i Then
'                                        If Datos.Fields(i) = vbNewLine Then
'                                            texto = texto & vbNewLine
'                                        Else
'                                            texto = texto & ArmarDato("DOC", "Tail" & CStr(ENTRAR), CStr(i), Datos.Fields(i)) & IIf(Separador = "tab", vbTab, Separador)
'                                        End If
'                                    ElseIf Datos.Fields.Count - 1 = i Then
'                                        If Datos.Fields(i) = vbNewLine Then
'                                            texto = texto & vbNewLine
'                                        Else
'                                            texto = texto & ArmarDato("DOC", "Tail" & CStr(ENTRAR), CStr(i), Datos.Fields(i))
'                                        End If
'                                    End If
'                                Next i
'                            End If
'                            Datos.MoveNext
'                        Loop
'                        ENTRAR = ENTRAR + 1
                    End If
                    Set Datos = Datos.NextRecordset()
                Loop
            End If
            VolcarTrama = PrefIni + texto + PrefFin
        End If
   VolcarTrama = Trim(VolcarTrama)
Exit Function
fin:
    VolcarTrama = "" & error
    Call Log_Inforest("Paperlees", "VolcarTrama", " Revsar SP TCPIP", "Caja: " + sCaja, NumTail, error, "", "FALLA PROCESO", sUsuario)
End Function
Public Function DesencriptaTrama(Datos As String, NumTail As String) As String()
Call DatosTrama("RECEIVE")
Dim Rs(1 To 1, 1 To 100) As String
Dim tOt, posi, cont As Integer
Dim Textoreceive As String
On Error GoTo fin

    Datos = Replace(Datos, IIf(PrefIni = "STX", "", PrefIni), "")
    Datos = Replace(Datos, IIf(PrefFin = "ETX", "", PrefIni), "")
    tOt = Len(Datos)
    cont = 0
    '------------------------------------- Captura y separa los valores de la trama de respuesta------------------------------------------------
    Do While tOt <> 0
        cont = cont + 1
        posi = InStr(1, Datos, IIf(Separador = "tab", vbTab, Separador))
        If posi > 0 Then
            Textoreceive = Mid(Datos, 1, posi - 1)
            Rs(1, cont) = Textoreceive
            Datos = Mid(Datos, posi + 1, tOt)
        Else
            Textoreceive = Mid(Datos, 1, tOt)
            Rs(1, cont) = Textoreceive
            Datos = ""
        End If
        tOt = Len(Datos)
    Loop
    '--------------------------------------------------Header-------------------------------------------------------------------------
    Dim iii As Integer
    Dim Res As String
    For iii = 0 To cantHeader - 1
        Res = ArmarDatoReceive("RECEIVE", "Header", CStr(iii), Rs(1, iii + 1))
        If Res = "EI0001" Then
            GoTo fin
        End If
    Next iii
    Dim contw As Integer
    contw = 0
    For iii = cantHeader To cont - 1
        Res = ArmarDatoReceive("RECEIVE", "Tail" & NumTail, CStr(contw), Rs(1, iii + 1))
        If Res = "EI0001" Then
            GoTo fin
        End If
        contw = contw + 1
    Next iii
    
    '----------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------Tail---------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------------------------------------------------------
    DesencriptaTrama = Rs
Exit Function
fin:
    Rs(1, 1) = "EI0001"
    Rs(1, 2) = IIf(error = "", "La trama de respuesta invalida Segun configuracion!. Revisar Configuracion", error)
    DesencriptaTrama = Rs
    Call Log_Inforest("Paperlees", "DesencriptaTrama", Datos, "Caja: " + sCaja, NumTail, error, "", "FALLA DESENCRIPTA", sUsuario)
End Function
Private Function ArmarDatoReceive(TipoIni As String, HeaderTail As String, Item As String, dato As String) As String
On Error GoTo fin
    Dim Large, posicion, tamaño As Integer
    Dim DataType As String
    Dim Valor As String
    ArmarDatoReceive = dato
    
    Valor = (Trim(LeerIniTrama(App.Path + "\ConfiguracionTrama\Trama\" & TipoIni & ".INI", HeaderTail, "ITEM" & Item, ".")))

    posicion = InStr(1, Valor, "|")
    tamaño = Len(Trim(Valor))

    DataType = CStr(Mid(Valor, 1, posicion - 1))
    Large = CInt(Mid(Valor, posicion + 1, tamaño))

    If DataType = "a" Then ' caracter alfabetico
        If ValAlfabetico(dato) = "" Then
            GoTo fin:
        End If
    ElseIf DataType = "n" Then ' caracter numerico
        If Not IsNumeric(dato) Then
            GoTo fin:
        End If
    ElseIf DataType = "an" Then ' caracter alfanumerico

    Else ' no se encontro ningun caracter
        GoTo fin:
    End If

    If ArmarDatoReceive <> "" Then
        If Large > 0 Then
            If Len(dato) <> Large Then
                GoTo fin:
            End If
        End If
    End If

    Exit Function
fin:
    Call Log_Inforest("Paperlees", "ArmarDatoReceive", TipoIni, "Caja: " + sCaja, "", error, "", "FALLA TCP IP", sUsuario)
    ArmarDatoReceive = "EI0001"
End Function
Private Function LeerIniTrama(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
On Error GoTo fin:
    'Los parámetros son:
    'lpFileName:    La Aplicación (fichero INI)
    'lpAppName:     La sección que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolverá
    '               si no se encuentra la clave.
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vacía
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If
    sRetVal = String$(255, 0)
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIniTrama = lpString
    Else
        LeerIniTrama = Left(sRetVal, LTmp)
    End If
    Exit Function
fin:
    LeerIniTrama = ""
    Call Log_Inforest("Paperlees", "LeerIniTrama", lpFileName, "Caja: " + sCaja, lpAppName, error, lpKeyName, "FALLA LEER INI", sUsuario)
End Function
