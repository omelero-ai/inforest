VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "License"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public conexion         As String
Public buffer           As String
Public l                As Long
Public relleno          As String
Public random           As String
Public hkestado         As String
Public clave1           As String
Public clave2           As String

Public Sub Class_Initialize()
    InitSB
End Sub

Private Sub InitBuffer()
'Inicializa con 10 bytes al azar la string usada para
'pasar los parámetros a la función HARDkey()
    
    Randomize
    relleno = Space(200)
    buffer = ""
    For l = 1 To 10
        buffer = buffer + Chr(Int((Rnd * 255) + 1))
    Next l
End Sub

Private Sub EncriptaString(buffer As String, password As String)
''Esta rutina encripta la cadena que se pasa como parámetro
''a la función HARDkey().
'
'    Dim i           As Integer
'    Dim ctemp       As Integer
'    Dim cAnterior   As Integer
'    Dim k           As Integer
'    Dim pw          As Integer
'    Dim bufEnc      As String
'
'    cAnterior = 0
'    bufEnc = ""
'    For i = 0 To 199
'        ctemp = Asc(Mid(buffer, i + 1, 1))
'        If (ctemp < 0) Then
'            ctemp = ctemp + 256
'        End If
'        ctemp = ctemp Xor sBox1(cAnterior)
'        For k = 0 To 15
'            pw = Asc(Mid(password, k + 1, 1))
'            If ((k Mod 2) = 1) Then
'                ctemp = ctemp Xor sBox1(sBox2(pw))
'                ctemp = sBox2(ctemp)
'            Else
'                ctemp = ctemp Xor sBox2(sBox1(pw))
'                ctemp = sBox1(ctemp)
'            End If
'        Next k
'        ctemp = ctemp Xor sBox1(i)
'        cAnterior = ctemp
'        bufEnc = bufEnc + Chr(ctemp)
'    Next i
'    buffer = bufEnc

'Esta rutina encripta la cadena que se pasa como parámetro
'a la función HARDkey().

    Dim i           As Integer
    Dim ctemp       As Integer
    Dim cAnterior   As Integer
    Dim k           As Integer
    Dim pw          As Integer
    Dim bufEnc      As String
    
    cAnterior = 0
    bufEnc = ""
    For i = 0 To 199
        ctemp = Asc(Mid(buffer, i + 1, 1))
        If (ctemp < 0) Then
            ctemp = ctemp + 256
        End If
        ctemp = ctemp Xor sBox1(cAnterior)
        For k = 0 To 15
            pw = Asc(Mid(password, k + 1, 1))
            If ((k Mod 2) = 1) Then
                ctemp = ctemp Xor sBox1(sBox2(pw))
                ctemp = sBox2(ctemp)
            Else
                ctemp = ctemp Xor sBox2(sBox1(pw))
                ctemp = sBox1(ctemp)
            End If
        Next k
        ctemp = ctemp Xor sBox1(i)
        cAnterior = ctemp
        bufEnc = bufEnc + Chr(ctemp)
    Next i
    buffer = bufEnc
    
End Sub

Private Sub DesencriptaString(buffer As String, password As String)
''Esta rutina desencripta la cadena que devuelve la
''función HARDkey().
'
'    Dim i           As Integer
'    Dim ctemp       As Integer
'    Dim cAnterior   As Integer
'    Dim k           As Integer
'    Dim pw          As Integer
'    Dim bufEnc      As String
'
'    cAnterior = 0
'    bufEnc = ""
'    For i = 0 To 199
'        ctemp = Asc(Mid(buffer, i + 1, 1))
'        If (ctemp < 0) Then
'            ctemp = ctemp + 256
'        End If
'        ctemp = ctemp Xor sBox1(cAnterior)
'        For k = 0 To 15
'            pw = Asc(Mid(password, k + 1, 1))
'            If ((k Mod 2) = 1) Then
'                ctemp = ctemp Xor sBox1(sBox2(pw))
'                ctemp = sBox2(ctemp)
'            Else
'                ctemp = ctemp Xor sBox2(sBox1(pw))
'                ctemp = sBox1(ctemp)
'            End If
'        Next k
'        ctemp = ctemp Xor sBox1(i)
'        cAnterior = Asc(Mid(buffer, i + 1, 1))
'        bufEnc = bufEnc + Chr(ctemp)
'    Next i
'    buffer = bufEnc
'Esta rutina desencripta la cadena que devuelve la
'función HARDkey().
    
    Dim i           As Integer
    Dim ctemp       As Integer
    Dim cAnterior   As Integer
    Dim k           As Integer
    Dim pw          As Integer
    Dim bufEnc      As String
    
    cAnterior = 0
    bufEnc = ""
    For i = 0 To 199
        ctemp = Asc(Mid(buffer, i + 1, 1))
        If (ctemp < 0) Then
            ctemp = ctemp + 256
        End If
        ctemp = ctemp Xor sBox1(cAnterior)
        For k = 0 To 15
            pw = Asc(Mid(password, k + 1, 1))
            If ((k Mod 2) = 1) Then
                ctemp = ctemp Xor sBox1(sBox2(pw))
                ctemp = sBox2(ctemp)
            Else
                ctemp = ctemp Xor sBox2(sBox1(pw))
                ctemp = sBox1(ctemp)
            End If
        Next k
        ctemp = ctemp Xor sBox1(i)
        cAnterior = Asc(Mid(buffer, i + 1, 1))
        bufEnc = bufEnc + Chr(ctemp)
    Next i
    buffer = bufEnc
End Sub

Private Function ValidaString(buffer As String, random As String) As Boolean
''Esta rutina analiza que la string devuelta por la
''función HARDkey() sea consistente.
'
'    Dim i      As Integer
'    Dim result As Boolean
'    Dim st     As String
'
'    For i = 1 To 10
'        Mid(buffer, i, 1) = Chr(sBox2(Asc(Mid(buffer, i, 1))))
'    Next i
'    result = True
'    For i = 1 To 10
'        If (Mid(buffer, i, 1) <> Mid(random, i, 1)) Then
'            result = False
'        End If
'    Next i
'    If (Mid(buffer, 11, 1) <> " ") Then
'        result = False
'    End If
'    If (Mid(buffer, 20, 1) <> " ") Then
'        result = False
'    End If
'    If (Mid(buffer, 26, 1) <> " ") Then
'        result = False
'    End If
'    If (Mid(buffer, 31, 1) <> "-") Then
'        result = False
'    End If
'    If (Val(Mid(buffer, 21, 5)) <> 0) Then
'        result = False
'    End If
'    st = Mid(buffer, 21, 5)
'    If (st = "00000") Then hkestado = "Se ha encontrado el HARDkey : " '""
'    If (st = "00002") Then hkestado = "No se encuentra dispositivo de seguridad HardKey en el servidor."
'    If (st = "00004") Then hkestado = "Formato de cadena o parámetro incorrecto"
'    If (st = "00010") Then hkestado = "Número de conexión no válida"
'    If (st = "00011") Then hkestado = "Usted ha superado la cantidad de licencias para este módulo"
'    If (st = "00012") Then hkestado = "Módulo yá en uso por la aplicación"
'    If (st = "00013") Then hkestado = "Módulo no levantado por la aplicación"
'    If (st = "00020") Then hkestado = "No hay drivers HARDkey instalados"
'    If (st = "00021") Then hkestado = "Versión de drivers obsoleta"
'    If (st = "00022") Then hkestado = "No hay drivers SuperPro instalados"
'    ValidaString = result
'Esta rutina analiza que la string devuelta por la
'función HARDkey() sea consistente.

    Dim i      As Integer
    Dim result As Boolean
    Dim st     As String
    
    For i = 1 To 10
        Mid(buffer, i, 1) = Chr(sBox2(Asc(Mid(buffer, i, 1))))
    Next i
    result = True
    For i = 1 To 10
        If (Mid(buffer, i, 1) <> Mid(random, i, 1)) Then
            result = False
        End If
    Next i
    If (Mid(buffer, 11, 1) <> " ") Then
        result = False
    End If
    If (Mid(buffer, 20, 1) <> " ") Then
        result = False
    End If
    If (Mid(buffer, 26, 1) <> " ") Then
        result = False
    End If
    If (Mid(buffer, 31, 1) <> "-") Then
        result = False
    End If
    If (Val(Mid(buffer, 21, 5)) <> 0) Then
        result = False
    End If
    st = Mid(buffer, 21, 5)
    If (st = "00000") Then hkestado = ""
    If (st = "00002") Then hkestado = "No se encontró protector"
    If (st = "00004") Then hkestado = "Formato de cadena o parámetro incorrecto"
    If (st = "00010") Then hkestado = "Número de conexión no válida"
    If (st = "00011") Then hkestado = "Se superó límite de usuarios permitidos"
    If (st = "00012") Then hkestado = "Módulo yá en uso por la aplicación"
    If (st = "00013") Then hkestado = "Módulo no levantado por la aplicación"
    If (st = "00020") Then hkestado = "No hay drivers HARDkey instalados"
    If (st = "00021") Then hkestado = "Versión de drivers obsoleta"
    If (st = "00022") Then hkestado = "No hay drivers SuperPro instalados"
    ValidaString = result
End Function

Public Sub SetClaves(ByVal pClave1 As String, ByVal pClave2 As String)
'    clave1 = Hex$(pClave1 - 500)
'    clave1 = IIf(Len(clave1) < 5, "0" + clave1, clave1)
'
'    clave2 = Hex$(pClave2 - 500)
'    clave2 = IIf(Len(clave2) < 5, "0" + clave2, clave2)
        clave1 = pClave1
        clave2 = pClave2
End Sub

Public Function IniciaConexion(ByVal pAplicacionId As Integer) As String
'    'Se fuerza a realizar una busqueda localmente para evitar que no se descuente la licencia
'    'correspondiente si el HardKey se encuentra primero localmente
'    InitBuffer
'    random = buffer
'    buffer = buffer + " 00000000 00000 00000 0009 00101 00000 00000 00000 00000 00000 00000 00000"
'    buffer = buffer + relleno
'    Call EncriptaString(buffer, password)
'    l = HARDkey(buffer)
'    '-----------------------------------------------------------------------------------------
'
'    Dim appId As String
'    appId = IIf(pAplicacionId < 10, "0" + CStr(pAplicacionId), CStr(pAplicacionId))
'
'    InitBuffer
'    random = buffer
'    buffer = buffer + " 00000000 " + clave1 + " " + clave2 + " 0000 00 0" ' + appId + " 3"
'    buffer = buffer + relleno
'    Call EncriptaString(buffer, password)
'    l = HARDkey(buffer)
'    Call DesencriptaString(buffer, password)
    
    
    'Este botón chequea la presencia de la llave invocando
'el comando Iniciar Conexión.

    'Primero deshabilito el chequeo en red, limitando la
    'busqueda de la llave a los puertos paralelo locales.
    'Para chequear en red y local elimine esta parte.
    InitBuffer
    'Selección de distintos chequeos
    ' Los siguiente ejemplos muestran como armar la string para los distintos
    ' tipos de chequeos, según se desee buscar la llave en los puertos locales
    ' o en la red, y el orden en que se realizará la búsqueda.
    '    Chequeo de la llave SuperPro en puertos USB y paralelo.
    '          strIN := strIN + '00004 00000 00000 00000 00000 00000 00000 00000';
    '    Chequeo de la llave SuperPro y HARDkey en puertos USB y paralelo.
    '    En este caso primero buscará la llave HARDkey y luego las llaves SuperPro
    '          strIN := strIN + '00001 00004 00000 00000 00000 00000 00000 00000';
    '    Chequeo de la llave HARDkey en red. En este caso busca primero la llave
    '    por MAILSLOT, luego por el protocolo IPX
    '          strIN := strIN + '00002 00003 00000 00000 00000 00000 00000 00000';
    buffer = buffer + " 00000000 00000 00000 0009 00001 00000 00000 00000 00000 00000 00000 00000"
    buffer = buffer + relleno
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)

    'Chequeo la llave invocando el comando Iniciar Conexión
    InitBuffer
    random = buffer
    buffer = buffer + " 00000000 " + clave1 + " " + clave2 + " 0000 00 0"
    buffer = buffer + relleno
    'CadenaEnviada.Text = Mid(buffer, 12)
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    
    'Analizo la respuesta para determinar si está la llave
    If (ValidaString(buffer, random) = True) Then
        conexion = Mid(buffer, 12, 8)
    Else
        conexion = "00000000"
    End If
    
    IniciaConexion = hkestado
End Function

Public Function VerificaConexion() As Boolean
    Dim verificacion As Boolean
    verificacion = False
    
    InitBuffer
    
    random = buffer
    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0001"
    buffer = buffer + relleno
    
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    'Analizo respuesta para ver si el comando fue completado
    If (ValidaString(buffer, random) = True) Then
        verificacion = True
    Else
        verificacion = False
    End If
    
    If verificacion = False Then
    
    Dim i As Integer
    
        For i = 0 To 2
        If verificacion = False Then
            InitBuffer
            
            random = buffer
            buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0001"
            buffer = buffer + relleno
            
            Call EncriptaString(buffer, password)
            l = HARDkey(buffer)
            Call DesencriptaString(buffer, password)
            
            'Analizo respuesta para ver si el comando fue completado
            If (ValidaString(buffer, random) = True) Then
                verificacion = True
            Else
                verificacion = False
            End If
        End If
        Next i
    End If
    
    VerificaConexion = verificacion
End Function

Private Function PISaldo(ByVal fVencimiento As Date, ByVal fRTC As Date) As String
    Dim result As Integer
    result = DateDiff("d", fRTC, fVencimiento)
    
    Dim mensajeRetorno As String
    
    If result > 5 Then
        mensajeRetorno = "2"
    End If
    
    If result <= 5 And result > 0 Then
        mensajeRetorno = "Faltan " + CStr(result) + " dia(s) para que caduque su licencia."
    End If
    
    If result <= 0 Then
        mensajeRetorno = "4"
    End If
    
    PISaldo = mensajeRetorno
End Function

Private Function PIChequeDif(ByVal fVencimiento As Date, ByVal fRTC As Date) As String
    Dim result As Integer
    result = DateDiff("d", fRTC, fVencimiento)
    
    Dim mensajeRetorno As String
    
    If result > 4 Then
        mensajeRetorno = "2"
    End If
    
    If result <= 4 And result > 0 Then
        mensajeRetorno = "La letra/cuota vence en " + CStr(result) + " dia(s)."
    End If
    
    If result = 0 Then
        mensajeRetorno = "La letra/cuota ha vencido, le quedan 8 dia(s)."
    End If
    
    If result < 0 And result > -8 Then
        result = 3 + result
        mensajeRetorno = "La letra/cuota ha vencido, le quedan " + CStr(result) + " dia(s)."
    End If
    
    If result <= -8 Then
        mensajeRetorno = "4"
    End If
    
    PIChequeDif = mensajeRetorno
End Function

Private Function PILetras(ByVal fVencimiento As Date, ByVal fRTC As Date) As String
    Dim result As Integer
    result = DateDiff("d", fRTC, fVencimiento)
    
    Dim mensajeRetorno As String
    
     If result > 4 Then
        mensajeRetorno = "2"
    End If
    
    If result <= 4 And result > 0 Then
        mensajeRetorno = "La letra/cuota vence en " + CStr(result) + " dia(s)."
    End If
    
    If result = 0 Then
        mensajeRetorno = "La letra/cuota ha vencido, le quedan 8 dia(s)."
    End If
    
    If result < 0 And result > -8 Then
        result = 8 + result
        mensajeRetorno = "La letra/cuota ha vencido, le quedan " + CStr(result) + " dia(s)."
    End If
    
    If result <= -8 Then
        mensajeRetorno = "4"
    End If
    
    PILetras = mensajeRetorno
End Function

Public Function PIPostVenta() As String
    
    Dim fechaRTC As Date
    fechaRTC = ObtenerRTC
    
    Dim fechaVencimientoMant As Date
    Dim fVencimientoMant As String
    fVencimientoMant = LeeCadena("00073", 2) + "/" + LeeCadena("00075", 2) + "/20" + LeeCadena("00077", 2)
    fechaVencimientoMant = CDate(Format(fVencimientoMant, "dd mm yyyy"))

    Dim result As Integer
    result = DateDiff("d", fechaRTC, fechaVencimientoMant)
    
    Dim mensajeRetorno As String
    If result > 0 Then
        mensajeRetorno = "Usted cuenta con " + CStr(result) + "/30 dia(s) de soporte técnico."
    Else
        mensajeRetorno = "3"
    End If
    PIPostVenta = mensajeRetorno
End Function

Private Function ObtenerFechaVencimiento() As Date
    'Se lee la celda de Dia/Mes
    Dim diaMes As String
    diaMes = LeerCelda(Nro04)
    diaMes = IIf(Len(diaMes) = 3, "0" + diaMes, diaMes)
    
    'Se lee la celda de Año
    Dim anio As String
    anio = LeerCelda(Nro05)
    
    Dim dia As String
    Dim mes As String
    
    dia = Mid(diaMes, 1, 2)
    mes = Mid(diaMes, 3, 2)
    
    Dim fechaVencimiento As Date
    fechaVencimiento = CDate(Format(dia + "/" + mes + "/" + anio, "dd mm yyyy"))
    
    ObtenerFechaVencimiento = fechaVencimiento
End Function

Private Function ObtenerRTC() As Date
    InitBuffer
    random = buffer
    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0010"
    buffer = buffer + relleno
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    Dim strFechaRTC As String
    strFechaRTC = Mid(buffer, 12)
    strFechaRTC = Mid(strFechaRTC, 27, 8)
    
    Dim fechaRTC As Date
    fechaRTC = CDate(Format(strFechaRTC, "dd mm yyyy"))
    
    ObtenerRTC = fechaRTC
End Function

Public Function ValidaFechaVencimiento() As String
    'Si la funcion retorna "1" quiere decir que el usuario tiene licencia de por vida
    'Si la funcion retorna "2" quiere decir que aun le queda tiempo suficiente antes que caduque su licencia
    'Si la funcion retorna "Faltan x dia(s) para que caduque su licencia." quiere decir que ya esta a pocos dias que caduque la licencia
    'Si la funcion retorna "4" quiere decir que la licencia caduco
    'Si la funcion retorna "5" quiere decir que no se encontro la llave conectada
    
    If Not VerificaConexion Then
        ValidaFechaVencimiento = "5"
        Exit Function
    End If
    
    Dim resultado As String
    Dim fechaVencimiento As Date
    Dim fechaRTC As Date
    
    fechaRTC = ObtenerRTC
    fechaVencimiento = ObtenerFechaVencimiento

    Dim anio As String
    anio = Format(fechaVencimiento, "yyyy")

    If anio = "1970" Then
        'Tiene licencia de por vida
        ValidaFechaVencimiento = "1"
        Exit Function
    End If

    Dim tipoContrato As Long
    tipoContrato = CLng(LeeCadena("00068", 1))

    Select Case tipoContrato
    
    Case contrato.PI_Saldo
        resultado = PISaldo(fechaVencimiento, fechaRTC)
    Case contrato.PI_ChequeDif
        resultado = PIChequeDif(fechaVencimiento, fechaRTC)
    Case contrato.PI_Letras
        resultado = PILetras(fechaVencimiento, fechaRTC)
    End Select
    
    ValidaFechaVencimiento = resultado
End Function

Public Function FinalizarConexion(ByVal pAplicacionId As Integer) As Boolean
'    Dim appId As String
'    appId = IIf(pAplicacionId < 10, "0" + CStr(pAplicacionId), CStr(pAplicacionId))
'
'    InitBuffer
'    random = buffer
'    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0002 " + appId
'    buffer = buffer + relleno
'
'    Call EncriptaString(buffer, password)
'    l = HARDkey(buffer)
'    Call DesencriptaString(buffer, password)
'
'    If (ValidaString(buffer, random) = True) Then
'        FinalizarConexion = True
'    Else
'        FinalizarConexion = False
'    End If
    InitBuffer
    random = buffer
    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0002 00"
    buffer = buffer + relleno
    'CadenaEnviada.Text = Mid(buffer, 12)
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    'Analizo respuesta para ver si el comando fue completado
    'RespuestaRecibida.Text = Mid(buffer, 12)
    If (ValidaString(buffer, random) = True) Then
        FinalizarConexion = True
        'LeerCadena.Enabled = False
        'FinConexion.Enabled = False
        'Estado.Text = "Conexión finalizada con exito"
    Else
        'Estado.Text = hkestado
        FinalizarConexion = False
    End If
End Function

Private Function LeerCelda(ByVal nroCelda As Celda) As String

    InitBuffer
    random = buffer
    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0003 " + IIf(nroCelda < 10, "0000" + CStr(nroCelda), "000" + CStr(nroCelda))
    buffer = buffer + relleno
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
     
    Dim str As String
    str = Mid(buffer, 38)
    
    LeerCelda = DecimalToHexadecimal(str)
End Function

Private Function LeeCadena(ByVal inicio As String, ByVal longitud As Integer) As String
    
'    Dim lenght As String
'    lenght = IIf(longitud < 10, "0000" + CStr(longitud), "000" + CStr(longitud))
'
'    InitBuffer
'    random = buffer
'    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0005 " + inicio + " " + lenght
'    buffer = buffer + relleno
'    Call EncriptaString(buffer, password)
'    l = HARDkey(buffer)
'    Call DesencriptaString(buffer, password)
'
'    LeeCadena = Mid(buffer, 38, longitud)

'Este botón leer los 64 bytes de memoria no-volátil
'de la llave reservados para el desarrollador

    InitBuffer
    random = buffer
    buffer = buffer + " " + conexion + " 00000 00000 0005 00044 00064"
    buffer = buffer + relleno
    'CadenaEnviada.Text = Mid(buffer, 12)
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    'Analizo respuesta para ver si el comando fue completado
    'RespuestaRecibida.Text = Mid(buffer, 12)
    If (ValidaString(buffer, random) = True) Then
        'Estado.Text = "Cadena leida con exito"
    Else
        'Estado.Text = hkestado
    End If
    LeeCadena = ""
End Function

Public Function ValidaLlave() As Boolean

'    If Not VerificaConexion Then
'        ValidaLlave = False
'        Exit Function
'    End If
'
'    InitBuffer
'    random = buffer
'    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0010"
'    buffer = buffer + relleno
'    Call EncriptaString(buffer, password)
'    l = HARDkey(buffer)
'    Call DesencriptaString(buffer, password)
'
'    Dim rtc As String
'    rtc = Mid(buffer, 12)
'    rtc = Mid(rtc, 39, 2)
'
'    Dim rango As Integer
'    rango = CInt(rtc)
'
'    Dim resultado As String
'    Dim value As Boolean
'    value = False
'
'    If rango >= 0 And rango <= 15 Then
'        resultado = LeeCadena("00044", 6)
'
'        If resultado = "015001" Then
'            value = True
'        End If
'    End If
'
'    If rango >= 16 And rango <= 30 Then
'        resultado = LeeCadena("00050", 6)
'
'        If resultado = "030162" Then
'            value = True
'        End If
'    End If
'
'    If rango >= 31 And rango <= 45 Then
'        resultado = LeeCadena("00056", 6)
'
'        If resultado = "045313" Then
'            value = True
'        End If
'    End If
'
'    If rango >= 46 And rango <= 59 Then
'        resultado = LeeCadena("00062", 6)
'
'        If resultado = "059464" Then
'            value = True
'        End If
'    End If
    
    ValidaLlave = True
End Function

Private Function DecimalToHexadecimal(ByVal str As String) As String
    Dim resultado As String
    resultado = Hex$(str)
    
    DecimalToHexadecimal = resultado
End Function

Public Function InicializaValores(ByVal password As String, ByVal tipoContrato As contrato) As Boolean
    If Not VerificaConexion Then
        InicializaValores = False
        Exit Function
    End If
    
    If password <> "1nfh0t3l" Then
        InicializaValores = False
        Exit Function
    End If

    If Not GrabaCadena(Byte_Nro44, 24, "015001030162045313059464") Then
        InicializaValores = False
        Exit Function
    End If
    
    If Not GrabaCadena(Byte_Nro68, 1, CStr(tipoContrato)) Then
        InicializaValores = False
        Exit Function
    End If
    
    InicializaValores = True
End Function

Private Function GrabaCadena(ByVal nroByte As Bite, ByVal longitud As Integer, ByVal cadena As String) As Boolean
    Dim lenght As String
    Dim bit As String
    
    lenght = IIf(longitud < 10, "0000" + CStr(longitud), "000" + CStr(longitud))
    bit = IIf(nroByte < 100, "000" + CStr(nroByte), "00" + CStr(nroByte))

    InitBuffer
    random = buffer
    buffer = buffer + " " + conexion + " " + clave1 + " " + clave2 + " 0006 " + bit + " " + lenght + " " + cadena

    buffer = buffer + relleno
    Call EncriptaString(buffer, password)
    l = HARDkey(buffer)
    Call DesencriptaString(buffer, password)
    
    If (ValidaString(buffer, random) = True) Then
        GrabaCadena = True
    Else
        GrabaCadena = False
    End If
End Function

