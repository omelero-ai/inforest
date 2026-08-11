Attribute VB_Name = "modImpresoraFiscal"
Option Explicit

Function SeteoFactura(PrinterOCX As PrinterFiscal) As Boolean
' **************************************************************************
' Propiedad intelectual EPSON ARGENTINA S.A.
' Programador: Gomez Guillermo
' Este software se entrega con fines didácticos y sin garantia alguna.
' EPSON NO ASUME responsabilidad legal alguna. El programador usa esta información
' bajo su propio riesgo y responsabilidad.
' **************************************************************************

'Ejemplo de seteo de Zona, Headers y Trailers para una factura
'TAMAÑO  : Form.continuo 12'' x 10''
'IMPRESOR: LX-300 Fiscal, FX-880 Fiscal

'PrinterOCX : es el objeto PrinterFiscal (OCX) a travez del cual
'             se accede a la impresora fiscal

Dim Continuar As Boolean

'** DEFINICION GENERAL DEL FORMULARIO

    '** Configura el tamaño de la hoja, calculando las líneas para 1/6 de pulgada y
    '** las columnas en 12 caracteres por pulgada
    Continuar = PrinterOCX.SetPaperSize(90, 72)

    '** Esta instruccion configura la estacion para la emision Facturas A por Main Sheet
    '** no se debe eliminar
    If Continuar Then
        Continuar = PrinterOCX.SetPreference("P", "D", "M", "F", "A", "N")
    End If
    '** Esta instruccion configura la estacion para la emision de documento no fiscal
    '** no debe eliminarse, si es necesario se puede modificar
    If Continuar Then
        Continuar = PrinterOCX.SetPreference("P", "P", "R", "S", "U", "90", "72")
    End If
'** DEFINICION DE ZONAS
    '
    '** ZONA(010) Letra Comprobante
    If Continuar Then
        Continuar = PrinterOCX.SetZone(10, 43, 4, 44, 4)
    End If
    '** ZONA(022) Logo
    If Continuar Then
        Continuar = PrinterOCX.SetZone(22, 1, 4, 40, 4)
    End If
    '** ZONA(024) Fantasía y Razon social
    If Continuar Then
        Continuar = PrinterOCX.SetZone(24, 1, 5, 40, 7)
    End If
    '** ZONA(026) Datos Vendedor
    If Continuar Then
        Continuar = PrinterOCX.SetZone(26, 1, 9, 40, 14)
    End If
    '** ZONA(034) Nro Factura y fecha
    If Continuar Then
        Continuar = PrinterOCX.SetZone(34, 49, 5, 87, 7)
    End If
    '** ZONA(036) Cuit e impuestos
    If Continuar Then
        Continuar = PrinterOCX.SetZone(36, 45, 9, 87, 14)
    End If
    '** ZONA(042) Datos del comprador
    If Continuar Then
        Continuar = PrinterOCX.SetZone(42, 1, 16, 40, 20)
    End If
    '** ZONA(044) IVA Comprador
    If Continuar Then
        Continuar = PrinterOCX.SetZone(44, 1, 21, 40, 22)
    End If
    '** ZONA(052) CUIT Comprador
    If Continuar Then
        Continuar = PrinterOCX.SetZone(52, 45, 16, 87, 20)
    End If
    '** ZONA(054) Nro remitos
    If Continuar Then
        Continuar = PrinterOCX.SetZone(54, 45, 21, 87, 22)
    End If
    '** ZONA(060) Zona de detalle de venta
    If Continuar Then
        Continuar = PrinterOCX.SetZone(60, 1, 23, 87, 42)
    End If
    '** ZONA(061) Cantidad Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(61, 1, 23, 12, 42)
    End If
    '** ZONA(062) Descripción Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(62, 14, 23, 34, 42)
    End If
    '** ZONA(063) Precio Unitario Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(63, 36, 23, 48, 42)
    End If
    '** ZONA(064) Tasa de IVA Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(64, 50, 23, 59, 42)
    End If
    '** ZONA(065) Tasa Ajuste Variable Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(65, 61, 23, 70, 42)
    End If
    '** ZONA(066) Precio total Item
    If Continuar Then
        Continuar = PrinterOCX.SetZone(66, 72, 23, 87, 42)
    End If
    '** Esta instruccion configura el offset que se usará para imprimir los subtotales
    '** debe ser menor o igual a las filas de la zona 60, por omisión se toma
    '** las filas de la zona 60
    If Continuar Then
        Continuar = PrinterOCX.SetPreference("P", "C", "O", "901", "0", "20")
    End If
    '** ZONA(070) Mensajes de advertencia fiscal
    If Continuar Then
        Continuar = PrinterOCX.SetZone(70, 1, 23, 87, 42)
    End If
    '** ZONA(081) Descripción Percepcion de IVA RI
    If Continuar Then
        Continuar = PrinterOCX.SetZone(81, 33, 23, 41, 49)
    End If
    '** ZONA(082) Monto Percepcion de IVA RI
    If Continuar Then
        Continuar = PrinterOCX.SetZone(82, 43, 23, 54, 49)
    End If
    '** ZONA(083) Tasa de acrecentamiento RNI
    If Continuar Then
        Continuar = PrinterOCX.SetZone(83, 33, 23, 41, 49)
    End If
    '** ZONA(084) Monto de acrecentamiento RNI
    If Continuar Then
        Continuar = PrinterOCX.SetZone(84, 43, 23, 54, 49)
    End If
    '** ZONA(085) Descr.otras percep.
    If Continuar Then
        Continuar = PrinterOCX.SetZone(85, 56, 23, 62, 49)
    End If
    '** ZONA(086) Monto otras percep.
    If Continuar Then
        Continuar = PrinterOCX.SetZone(86, 64, 23, 75, 49)
    End If
    '** ZONA(090) Zona de subtotal y cierre
    If Continuar Then
        Continuar = PrinterOCX.SetZone(90, 1, 23, 87, 49)
    End If
    '** ZONA(091) Conceptos no gravados
    If Continuar Then
        Continuar = PrinterOCX.SetZone(91, 1, 23, 10, 49)
    End If
    '** ZONA(092) Tasa de IVA
    If Continuar Then
        Continuar = PrinterOCX.SetZone(92, 12, 23, 17, 49)
    End If
    '** ZONA(093) Monto de IVA
    If Continuar Then
        Continuar = PrinterOCX.SetZone(93, 19, 23, 31, 49)
    End If
    '** ZONA(094) Percepciones de IVA
    If Continuar Then
        Continuar = PrinterOCX.SetZone(94, 33, 23, 54, 49)
    End If
    '** ZONA(095) Otras percepciones
    If Continuar Then
        Continuar = PrinterOCX.SetZone(95, 56, 23, 75, 49)
    End If
    '** ZONA(097) Subtotal
    If Continuar Then
        Continuar = PrinterOCX.SetZone(97, 77, 23, 87, 49)
    End If
    '** ZONA(100) Zona de total documento
    If Continuar Then
        Continuar = PrinterOCX.SetZone(100, 1, 50, 87, 50)
    End If
    '** ZONA(101) Título de Total documento
    If Continuar Then
        Continuar = PrinterOCX.SetZone(101, 1, 50, 55, 50)
    End If
    '** ZONA(102) Monto de Total documento
    If Continuar Then
        Continuar = PrinterOCX.SetZone(102, 56, 50, 87, 50)
    End If
    '** ZONA(110) Zona de detalle de pagos
    If Continuar Then
        Continuar = PrinterOCX.SetZone(110, 1, 51, 87, 55)
    End If
    '** ZONA(111) Descripción de Pagos
    If Continuar Then
        Continuar = PrinterOCX.SetZone(111, 1, 51, 39, 55)
    End If
    '** ZONA(112) Monto de Pagos
    If Continuar Then
        Continuar = PrinterOCX.SetZone(112, 41, 51, 87, 55)
    End If
    '** ZONA(120) Leyendas de Pie de Factura
    If Continuar Then
        Continuar = PrinterOCX.SetZone(120, 1, 56, 87, 58)
    End If
    '** ZONA(142) Logo Fiscal
    If Continuar Then
        Continuar = PrinterOCX.SetZone(142, 1, 59, 40, 63)
    End If
    '** ZONA(144) Original / Copia
    If Continuar Then
        Continuar = PrinterOCX.SetZone(144, 45, 59, 87, 63)
    End If
'** DEFINICION DE HEADERS y TRAILERS
    '
    '** HEADER/TRAILER (01) Nombre de Fantasía - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 1, Chr(127))
    End If
    '** HEADER/TRAILER (02) Nombre de Fantasía - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 2, Chr(127))
    End If
    '** HEADER/TRAILER (03) Nombre de Fantasía - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 3, Chr(127))
    End If
    '** HEADER/TRAILER (04) Nombre de Fantasía - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 4, Chr(127))
    End If
    '** HEADER/TRAILER (05) Nombre de Fantasía - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 5, Chr(127))
    End If
    '** HEADER/TRAILER (11) Leyenda de pie 11
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 11, "Gracias por su compra")
    End If
    '** HEADER/TRAILER (12) Leyenda de pie 12 - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 12, Chr(127))
    End If
    '** HEADER/TRAILER (13) Leyenda de pie 13 - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 13, Chr(127))
    End If
    '** HEADER/TRAILER (14) Leyenda de pie 14 - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 14, Chr(127))
    End If
    '** HEADER/TRAILER (50) Domicilio Comercial
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 50, "Dato fijo")
    End If
    '** HEADER/TRAILER (51) Domicilio Comercial
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 51, "Capital Federal")
    End If
    '** HEADER/TRAILER (52) Domicilio Comercial - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 52, Chr(127))
    End If
    '** HEADER/TRAILER (53) Domicilio Fiscal
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 53, "Belgrano 951")
    End If
    '** HEADER/TRAILER (54) Domicilio Fiscal
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 54, "Capital Federal")
    End If
    '** HEADER/TRAILER (55) Domicilio Fiscal - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 55, Chr(127))
    End If
    '** HEADER/TRAILER (57) Impuestos Vendedor
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 57, "Ingresos Brutos")
    End If
    '** HEADER/TRAILER (58) Impuestos Vendedor - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 58, Chr(127))
    End If
    '** HEADER/TRAILER (59) Impuestos Vendedor - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 59, Chr(127))
    End If
    '** HEADER/TRAILER (60) Impuestos Vendedor - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 60, Chr(127))
    End If
    '** HEADER/TRAILER (61) Impuestos Vendedor - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 61, Chr(127))
    End If
    '** HEADER/TRAILER (62) Fecha de inicio de actividades
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 62, "Inicio de actividades 01/01/2000")
    End If
    '** HEADER/TRAILER (63) Leyendas Comerciales - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 63, Chr(127))
    End If
    '** HEADER/TRAILER (64) Leyendas Comerciales - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 64, Chr(127))
    End If
    '** HEADER/TRAILER (65) Leyendas Comerciales - Vacio
    If Continuar Then
        Continuar = PrinterOCX.SetGetHeaderTrailer("S", 65, Chr(127))
    End If
    
SeteoFactura = Continuar

' **************************************************************************
' Este software se entrega con fines didácticos y sin garantia alguna.
' EPSON NO ASUME responsabilidad legal alguna. El programador usa esta información
' bajo su propio riesgo y responsabilidad.
' **************************************************************************

End Function

Function ImpresoraFiscalEpson(IFEpson As PrinterFiscal, Index As Integer, DatosDoc As Recordset) As Integer
    Dim respuesta As Integer
    respuesta = 0
On Error GoTo fin
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'CIERRE Z
            respuesta = CierreZ(IFEpson)
        Case 1 ' CIERRE X
            respuesta = CierreX(IFEpson)
        Case 2 ' TIKET  A,B
                respuesta = TicketFactura(IFEpson, DatosDoc)
        Case 3 ' NOTAS DE CREDITO A, B
            respuesta = TicketFacturaNC(IFEpson, DatosDoc)
        Case 4 ' Valida Status
            'Respuesta = StatusFiscal(IFEpson)
    End Select
    ImpresoraFiscalEpson = respuesta
    Screen.MousePointer = vbDefault
Exit Function
fin:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & error
    ImpresoraFiscalEpson = respuesta
End Function
Private Function StatusFiscal(IFEpson As PrinterFiscal) As Boolean
On Error GoTo fin
    Dim respuesta As Boolean
    respuesta = True
    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600
    
    If respuesta Then respuesta = IFEpson.Status("P")
    'If Respuesta Then Respuesta = IFEpson.CutPaper()
    StatusFiscal = respuesta
    Exit Function
fin:
    MsgBox "Error:" & error
    StatusFiscal = False
End Function
Private Function CierreX(IFEpson As PrinterFiscal) As Integer
    Dim respuesta As Boolean
    Dim VALORRespuesta As Integer
    respuesta = True
    VALORRespuesta = 0
On Error GoTo fin

    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600
    
    If respuesta Then respuesta = IFEpson.CloseJournal("X", "P")
    
    If respuesta Then VALORRespuesta = 1
    
    CierreX = VALORRespuesta
    Exit Function
fin:
    MsgBox "Error:" & error
    CierreX = VALORRespuesta
End Function
Private Function CierreZ(IFEpson As PrinterFiscal) As Integer
    Dim respuesta As Boolean
    Dim VALORRespuesta As Integer
    respuesta = True
    VALORRespuesta = 0
On Error GoTo fin

    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600
    
    If respuesta Then respuesta = IFEpson.CloseJournal("Z")
    
    If respuesta Then VALORRespuesta = 1
    
    CierreZ = VALORRespuesta
    Exit Function
fin:
    MsgBox "Error:" & error
    CierreZ = VALORRespuesta
End Function
Private Function TicketFactura(IFEpson As PrinterFiscal, RsDatos As Recordset) As Integer
    
    Dim respuesta As Boolean
    Dim RespVALOR As Integer
    respuesta = True
    RespVALOR = 0
    
On Error GoTo fin

    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600

    Dim vTipoResponsable As String
    Dim TDComprador As String
    Dim LetraDoc As String
    LetraDoc = Left(RsDatos!tDocumento, 1)
    vTipoResponsable = Calcular("select top 1 isnull(vTR,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn)
    TDComprador = Left(Calcular("select top 1 isnull(TipoIdentidad,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn), 6)
    
    If respuesta Then respuesta = IFEpson.OpenInvoice("T", "C", LetraDoc, "1", "P", "12", "I", vTipoResponsable, Left(RsDatos!Cliente, 40), IIf(Left(Trim(RsDatos!direccion), 40) = "", "-", Left(Trim(RsDatos!direccion), 40)), TDComprador, Replace(RsDatos!Ruc, "-", ""), "N", String(40, "-"), "147 Telefono gratuito CABA, Area de ", "Defensa y Proteccion al Consumidor", String(40, "-"), "Obs. " & Left(IIf(RsDatos!tObservacion = "", Right(RsDatos!tDocumento, 8), RsDatos!tObservacion), 35), "G")
    
    If respuesta Then RespVALOR = 1
    
    Do While Not RsDatos.EOF
        If respuesta Then respuesta = IFEpson.SendInvoiceItem(Left(RsDatos!Producto, 20), Replace(Format(RsDatos!nCantidad, "0####.##0"), ".", ""), IIf(LetraDoc = "B", Replace(Format(RsDatos!nPrecioOficial, "0######.#0"), ".", ""), Replace(Format((RsDatos!nprecioVenta / (1 + (nPorcentaje1 + nPorcentaje2 + nPorcentaje3) / 100) + RsDatos!descunitario / RsDatos!nCantidad), "0######.#0"), ".", "")), Replace(Format(nPorcentaje1 * 100, "0##0"), ".", ""), "M", "0", "0", "", "", "", "1050", "0")
     RsDatos.MoveNext
    Loop
    RsDatos.MoveLast

    If respuesta Then respuesta = IFEpson.GetInvoiceSubtotal("P", "SUB TOTAL")
    
    If respuesta Then RespVALOR = 2
    
    If RsDatos!MotivoDescuento <> "" Then
        'ImprimeXLinea "Descuento     : " & RsImpresion!MotivoDescuento, 40, 16
        If RsDatos!nDescuento > 0 Then
            If respuesta Then respuesta = IFEpson.SendInvoicePayment(Left(RsDatos!MotivoDescuento, 22), Replace(Format(RsDatos!nDescuento, "0#########.#0"), ".", ""), "D")
        End If
    Else
        If RsDatos!nDescuento > 0 Then
            If respuesta Then respuesta = IFEpson.SendInvoicePayment("Descuento: ", Replace(Format(RsDatos!nDescuento, "0#########.#0"), ".", ""), "D")
        End If
    End If

    If lPagoAntesImpresion And RsDatos!nVenta > 0 Then
    Dim RsPagoDocumento As Recordset
    Dim i As Integer
    Set RsPagoDocumento = Lib.OpenRecordset("exec usp_Inforest_Impresion '" & RsDatos!tDocumento & "', 9 ", Cn)
        If RsPagoDocumento.RecordCount > 0 Then
            For i = 0 To RsPagoDocumento.RecordCount - 1
                If respuesta Then respuesta = IFEpson.SendInvoicePayment(RsPagoDocumento!pago + "(" + Trim(CStr(RsPagoDocumento!ValorReal)) + ")", Replace(Format(RsPagoDocumento!MN, "0#########.#0"), ".", ""), "T")
                RsPagoDocumento.MoveNext
            Next i
        End If
    End If
    
    If respuesta Then respuesta = IFEpson.CloseInvoice("T", LetraDoc, "FINALIZA")
    
    
    
    'If Respuesta Then Respuesta = IFEpson.CutPaper()
    TicketFactura = RespVALOR
    Exit Function
fin:
    MsgBox "Mensaje: " & error, vbInformation, "Facturacion impresora fiscal"
    TicketFactura = RespVALOR
End Function
Public Sub ImprimeFacturaElectronicaARGE(ByVal RsImpresion As Recordset, Optional CodigoHash As Image, Optional tTipoDoc As String, Optional Cabecera As Image, Optional pie As Image, Optional ByVal cadenaCodigoHash As String, Optional ByVal ImpDolarDelivery As Boolean, Optional ImageIntegracion As Image, Optional ImageMercadoPago As Image)
    On Error Resume Next
    Dim Item As String
    Dim xObservacion As String
    Dim sLinea As String
    Dim sLinea2 As String
    Dim txt As String
    Dim xRsPropiedad As ADODB.Recordset
    Dim xRsPropiedadCombo As ADODB.Recordset
    Dim nSumaNeto As Double
    Dim xTotalS As Double
    Dim xproducto As String
    Dim i As Integer
    Dim xExonerado As Boolean
    
    
    Dim xImpresionFE As String
    Dim xRsObtAumentaPrecio As ADODB.Recordset
    
    Dim xRsObtPrecioCombos As ADODB.Recordset
    Dim RsPagoDocumento As ADODB.Recordset
    Dim lcuentaCobrar As Boolean

    Dim nDescuentoNeto As String
    Dim IsqlTC As String
    Dim RsParametroTC As ADODB.Recordset
    
    'imprime tipo de cambio
    IsqlTC = "select lVerTCImp from TPARAMETRO"
    Set RsParametroTC = Lib.OpenRecordset(IsqlTC, Cn)
    '------------------------------------------
                        
    '---CESAR IMP
    Dim imp1 As Double
    Dim imp2 As Double
    Dim impxT As Double
    Dim impx1 As Double
    Dim impx2 As Double
    Dim impG As Double
    Dim xDescuento As Double
    Dim xRecargo As Double
    Dim xSubTotal As Double
    Dim xDiferencia As Double
    Dim xAumento As Double
    
    imp1 = Calcular("select ISNULL(Impuesto1,0) as Codigo from TPARAMETRO", Cn)
    imp2 = Calcular("select ISNULL(Impuesto2,0) as Codigo from TPARAMETRO", Cn)
    impxT = 1 + (imp1 / 100) + (imp2 / 100)
    impx1 = 1 + (imp1 / 100)
    impx2 = 1 + (imp2 / 100)
    impG = 1
    '------------------
    lImprimeImagCabDocumento = Calcular("select lImprimeImageCab as codigo from ttipodocumentoimpresora where ttipoemision='" & tTipoDoc & "' and  tcaja='" & sCaja & "'", Cn)
    lImprimeImagPieDocumento = Calcular("select lImprimeImagePie as codigo from ttipodocumentoimpresora where ttipoemision='" & tTipoDoc & "' and  tcaja='" & sCaja & "'", Cn)
    
    Dim ximprimeDetallado As String
    ximprimeDetallado = Calcular("select lImprimeresumen as codigo from ttipodocumentoimpresora where ttipoemision='" & tTipoDoc & "' and  tcaja='" & sCaja & "'", Cn)
    Dim CANT, valor As Integer
    Dim ProductoTem As String
    
    xTotalS = 0
    
    Dim xUnidadNegocio As String
    xUnidadNegocio = Calcular("SELECT tDetallado As Codigo FROM TTABLA WHERE TTABLA = 'UNIDADNEGOCIO' AND TCODIGO = (SELECT tUnidadNegocio FROM TCAJA WHERE TCAJA='" & RsImpresion!tCaja & "')", Cn)
      
  
    Isql = "SELECT tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TOPERADOR.tDetallado AS Operador, dbo.TPROPIEDAD.tDetallado AS Descripcion,dbo.TPRODUCTOPROPIEDAD.nCantidad " & _
           "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto INNER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "where tCodigoPedido='" & RsImpresion!tCodigoPedido & "' and TOPERADOR.lImprime=1 Order by dbo.TOPERADOR.nBoton"
    Set xRsPropiedad = Lib.OpenRecordset(Isql, Cn)
    Isql = "SELECT dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, dbo.TOPERADOR.tDetallado AS Operador, dbo.TPROPIEDAD.tDetallado AS Descripcion, dbo.TCOMBOPROPIEDAD.tCodigoPedido,dbo.TCOMBOPROPIEDAD.nCantidad " & _
           "FROM dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN dbo.TCOMBOPROPIEDAD ON dbo.TPROPIEDAD.tProducto = dbo.TCOMBOPROPIEDAD.tProducto AND dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TCOMBOPROPIEDAD.tCodigoPropiedad " & _
           "where tCodigoPedido='" & RsImpresion!tCodigoPedido & "'  and TOPERADOR.lImprime=1 Order by dbo.TOPERADOR.nBoton"
    Set xRsPropiedadCombo = Lib.OpenRecordset(Isql, Cn)
    Dim P As Integer
    Select Case pais
        Case "003" 'ARGENTINA
       
         
                    'Cabecera
                    If lImprimeImagCabDocumento Then
                                    If Cabecera <> 0 Then
                                    Printer.PaintPicture Cabecera, 500, Printer.CurrentY, 2860, 1600
                                    For P = 1 To 10
                                    Printer.Print
                                    Next P
                                    End If
                    End If

                    RsImpresion.MoveFirst
                    ImprimeXCentro sRazonSocial, 40
                    ImprimeXCentro sRazonComercial, 40
'                    sLinea = String((40 - Len(Trim("R.U.C.: " + sRUC))) / 2, " ") & "R.U.C.: " + sRUC
                    Select Case pais
                        Case "003" ' argentina
                            sLinea = String((40 - Len(Trim("C.U.I.T.: " + sRUC))) / 2, " ") & "C.U.I.T.: " + sRUC
                        Case Else ' PERU
                            sLinea = String((40 - Len(Trim("R.U.C.: " + sRUC))) / 2, " ") & "R.U.C.: " + sRUC
                    End Select
                    Printer.Print sLinea
                    ImprimeXCentro sDireccion, 40
                    If Trim(sDireccion2) <> "" Then
                        ImprimeXCentro sDireccion2, 40
                    End If
                    sLinea = String((40 - Len(Trim("Telefono: " + sTelefono))) / 2, " ") & "Telefono: " + sTelefono
                    Printer.Print sLinea
                    
                    If sFax <> "" Then
                        sLinea = String((40 - Len(Trim("Fax: " + sFax))) / 2, " ") & "Fax: " + sFax
                        Printer.Print sLinea
                    End If
                    
                    If tCodigoFE <> "999" Then
                        Printer.Print " "
                        If Calcular("select lidentidadruc  as Codigo  from vTipoDocumento where Codigo in(select tTipoDocumento from MDOCUMENTO where tDocumento ='" & RsImpresion!tDocumento & "')", Cn) = True Then
                            'ImprimeXCentro "FACTURA ELECTRONICA", 40
                            ImprimeXCentro "FACTURA A", 40
                        Else
                            'ImprimeXCentro "BOLETA DE VENTA ELECTRONICA", 40
                            ImprimeXCentro "FACTURA B", 40
                        End If
                        Printer.Print " "
                    End If
                    
                    ImprimeXCentro Mid(RsImpresion!tDocumento, 1, 1) & Mid(RsImpresion!tDocumento, 3, 4) & "-" & Mid(RsImpresion!tDocumento, 8, 8), 40
                    Printer.Print " "
                    If lImpClienteCabDoc Then
                        Printer.Print String(40, "-")
                        If Calcular("select upper(tResumido) as Codigo from vTipoIdentidad inner join TCLIENTE on vTipoIdentidad.Codigo=tcliente.tTipoIdentidad where tcliente.tCodigoCliente= (select isnull(tCodigoCliente,'') as codigo From MDOCUMENTO where tDocumento='" & RsImpresion!tDocumento & "')", Cn) = "DNI" Then
                            ImprimeXLinea "Cliente: " & RsImpresion!Cliente, 40, 9
                            Printer.Print "DNI    : " & RsImpresion!Ruc
                            ImprimeXLinea "Direcc.: " & RsImpresion!direccion, 40, 9
                        Else
                            ImprimeXLinea "Empresa: " & RsImpresion!Cliente, 40, 9
                            Printer.Print "RUC    : " & RsImpresion!Ruc
                            ImprimeXLinea "Direcc.: " & RsImpresion!direccion, 40, 9
                        End If
                        If lFEBiz Then
                            Dim tUbigeo As String
                            Dim sUbigeo As String
                            tUbigeo = Calcular("select top 1 isnull(tUbigeo,'') as Codigo from tcliente where tIdentidad='" & RsImpresion!Ruc & "'", Cn)
                            sUbigeo = Calcular("select isnull(tdistrito+'-'+tprovincia+'-'+tDepartamento,'') as codigo from tubigeo where tcodigo='" & tUbigeo & "'", Cn)
                            ImprimeXLinea "Ubigeo : " & tUbigeo & " " & sUbigeo, 40, 9
                        End If
                        Printer.Print String(40, "-")
                    End If
                    Printer.Print "Pedido        : " & RsImpresion!tCodigoPedido
                    Printer.Print "Fecha Emision : " & RsImpresion!fRegistro
                    Printer.Print "Tipo          : " & RsImpresion!TipoPedido
                    Printer.Print "Local         : " & xUnidadNegocio
                    Printer.Print "Caja          : " & RsImpresion!tCaja & " - " & RsImpresion!tUsuario
                    Printer.Print "Mesero        : " & RsImpresion!Mozo
                    If ImpDolarDelivery Then
                        Printer.Print "Tipo Moneda   : " & sMonedaE
                    Else
                        Printer.Print "Tipo Moneda   : " & sMonedaN
                    End If
                    
                    'tipo de cambio
                    If RsParametroTC!lVerTCImp = True Then
                    Printer.Print "T/Cambio  : " & Format(nTC, "###,##0.00")
                    End If
                    
                    Printer.Print "Mesa          : " & RsImpresion!Mesa
                    If lImprimeMotivoDescuentoFB And RsImpresion!MotivoDescuento <> "" Then
                        ImprimeXLinea "Descuento     : " & RsImpresion!MotivoDescuento, 40, 16
                    End If
                    If lObservacionCabDoc Then
                        ImprimeXLinea "Observacion   : " & RsImpresion!tObservacion, 40, 16
                    End If
                    Printer.Print " "
                    Printer.Print String(40, "-")
                    If Not IsNull(RsImpresion!lImpProdDesc) And (RsImpresion!lImpProdDesc) Then
                        Printer.Print "Codigo      Cant. UM   P.Unit      Total"
                        Printer.Print "Producto                          Dscto."
                    Else
                        Printer.Print "Producto      Cant.    P.Unit      Total"
                    End If
                    Printer.Print String(40, "-")
                    'Cuerpo
                    Item = ""
                    Do While Not RsImpresion.EOF
                    If ximprimeDetallado = False Then
                        xproducto = RsImpresion!Producto
                    Else
                        xproducto = RsImpresion!ProductoDetallado
                    End If
                    
                    If Calcular("select isnull(treserva,'') as Codigo from mdocumento where tdocumento='" & RsImpresion!tDocumento & "'", Cn) <> "" Then
                        xproducto = xproducto + ": " + Calcular("select isnull(tobservacion,'') as codigo from dpedido where tcodigopedido='" & RsImpresion!tCodigoPedido & "' and titem='" & RsImpresion!tItem & "'", Cn)
                    End If
                    
                       If Item <> RsImpresion!tCodigoPedido + RsImpresion!tItem Then
                          'Impresion de Propiedades
                          If lPropiedadDocumento Then
                             xRsPropiedad.Filter = "tCodigoPedido='" & Mid(Item, 1, 10) & "' and tItem='" & Mid(Item, 11, 3) & "'"
                             If Not xRsPropiedad.EOF Then
                                xRsPropiedad.MoveFirst
                                Do While Not xRsPropiedad.EOF
                                    'ImprimeXLinea "        * " & xRsPropiedad!Operador & " " & xRsPropiedad!Descripcion, 40, 8
                                    ImprimeXLinea "        * " & IIf(xRsPropiedad!nCantidad = 1, "", "(" & xRsPropiedad!nCantidad & ")") & xRsPropiedad!Operador & " " & xRsPropiedad!descripcion, 31, 10
                                   xRsPropiedad.MoveNext
                                Loop
                             End If
                                         
                             'Impresion de la Observacion
                             If lObservacionDocumento And Len(Trim(xObservacion)) > 0 Then
                                ImprimeXLinea "        - " & Trim(xObservacion), 33, 2
                             End If
                          End If
                          
                          '''''''''''''''''''''
                          If IsNull(RsImpresion!tOferta) Or RsImpresion!tOferta = "" Then
                            If Not (RsImpresion!lImpProdDesc) Then
                                CANT = Len(Trim(xproducto))
                                ProductoTem = xproducto
regreso4:
                            If (lMostrarPrecioItemsCombo) Then
                                 If (RsImpresion!nCombo >= 1) Then
                                    'GCAA 12082021
                                    'Isql = "select SUM(nAumento)total_descontar from tCombo where tCombo= '" & RsImpresion!tCodigoProducto & "' and tCodigoProducto in (select tProductoCombo  from CPEDIDO where tCodigoPedido = '" & RsImpresion!tCodigoPedido & "' and tItem = '" & RsImpresion!tItem & "' ) "
                                    Isql = " select isnull( sum(tCombo.nAumento*CPEDIDO.nCantidad),0) as total_descontar from CPEDIDO inner join tCombo on CPEDIDO.tProducto = tCombo.tCombo and CPEDIDO.tProductoCombo=tCombo.tCodigoProducto  where CPEDIDO.tCodigoPedido = '" & RsImpresion!tCodigoPedido & "' and CPEDIDO.tItem='" & RsImpresion!tItem & "' "
                                    Set xRsObtPrecioCombos = Lib.OpenRecordset(Isql, Cn)
                                    If valor = 0 Then
                                            If ImpDolarDelivery Then
                                                sLinea = Mid(xproducto & String(14, " "), 1, 14) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(7, " ") & Format(RsImpresion!nPrecioOficial / nTC, "###0.00"), 7) & " " & Right(String(10, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) / nTC, "###,##0.00"), 10)
                                            Else
                                                sLinea = Mid(xproducto & String(14, " "), 1, 14) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(7, " ") & Format(RsImpresion!nPrecioOficial - (xRsObtPrecioCombos!total_descontar), "###0.00"), 7) & " " & Right(String(10, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) - (xRsObtPrecioCombos!total_descontar * RsImpresion!nCombo), "###,##0.00"), 10)
                                            End If
                                            valor = 1
                                            CANT = Len(Mid(ProductoTem, 17))
                                            ProductoTem = Mid(ProductoTem, 17)
                                            Printer.Print sLinea
                                    Else
                                        sLinea = Mid(ProductoTem & String(16, " "), 1, 16) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(9, " ") & "           ", 9)
                                        CANT = Len(Mid(ProductoTem, 17))
                                        ProductoTem = Mid(ProductoTem, 17)
                                        Printer.Print sLinea
                                    End If
                                Else
                                    If valor = 0 Then
                                            If ImpDolarDelivery Then
                                                sLinea = Mid(xproducto & String(14, " "), 1, 14) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(7, " ") & Format(RsImpresion!nPrecioOficial / nTC, "###0.00"), 7) & " " & Right(String(10, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) / nTC, "###,##0.00"), 10)
                                            Else
                                                sLinea = Mid(xproducto & String(14, " "), 1, 14) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(7, " ") & Format(RsImpresion!nPrecioOficial, "###0.00"), 7) & " " & Right(String(10, " ") & Format(RsImpresion!nCantidad * RsImpresion!nPrecioOficial, "###,##0.00"), 10)
                                            End If
                                            valor = 1
                                            CANT = Len(Mid(ProductoTem, 17))
                                            ProductoTem = Mid(ProductoTem, 17)
                                            Printer.Print sLinea
                                    Else
                                        sLinea = Mid(ProductoTem & String(16, " "), 1, 16) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(9, " ") & "           ", 9)
                                        CANT = Len(Mid(ProductoTem, 17))
                                        ProductoTem = Mid(ProductoTem, 17)
                                        Printer.Print sLinea
                                    End If
                                End If
                            Else
                                If valor = 0 Then
                                    If ImpDolarDelivery Then
                                        sLinea = Mid(xproducto & String(11, " "), 1, 11) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(10, " ") & Format(IIf(Mid(RsImpresion!tDocumento, 1, 1) = "A", RsImpresion!nPrecioNeto, RsImpresion!nprecioVenta) / nTC, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format((RsImpresion!nCantidad * IIf(Mid(RsImpresion!tDocumento, 1, 1) = "A", RsImpresion!nPrecioNeto, RsImpresion!nprecioVenta)) / nTC, "###,##0.00"), 10)
                                    Else
                                        sLinea = Mid(xproducto & String(11, " "), 1, 11) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(10, " ") & Format(IIf(Mid(RsImpresion!tDocumento, 1, 1) = "A", RsImpresion!nPrecioNeto, RsImpresion!nprecioVenta), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(RsImpresion!nCantidad * IIf(Mid(RsImpresion!tDocumento, 1, 1) = "A", RsImpresion!nPrecioNeto, RsImpresion!nprecioVenta), "###,##0.00"), 10)
                                    End If
                                    valor = 1
                                    CANT = Len(Mid(ProductoTem, 17))
                                    ProductoTem = Mid(ProductoTem, 17)
                                    Printer.Print sLinea
                                Else
                                    sLinea = Mid(ProductoTem & String(16, " "), 1, 16) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(9, " ") & "           ", 9)
                                    CANT = Len(Mid(ProductoTem, 17))
                                    ProductoTem = Mid(ProductoTem, 17)
                                    Printer.Print sLinea
                                End If
                            End If
                                If CANT > 0 And ximprimeDetallado = True Then
                                GoTo regreso4
                                End If
                            End If
                            If (RsImpresion!lImpProdDesc) Then ' Para impresion con formato 2
                                If ImpDolarDelivery Then
                                    sLinea = Mid(RsImpresion!tCodigoProducto & String(8, " "), 1, 8) & " " & Right(String(7, " ") & Format(RsImpresion!nCantidad, "##0.00"), 7) & " " & Right(String(3, " ") & "UN", 3) & " " & Right(String(9, " ") & Format(RsImpresion!nPrecioOficial / nTC, "##0.00"), 9) & " " & Right(String(9, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) / nTC, "##,##0.00"), 9)
                                    sLinea2 = Mid(xproducto & String(30, " "), 1, 30) & " " & Right(String(9, " ") & Format((RsImpresion!descunitario) / nTC, "##,##0.00"), 9)
                                Else
                                    sLinea = Mid(RsImpresion!tCodigoProducto & String(8, " "), 1, 8) & " " & Right(String(7, " ") & Format(RsImpresion!nCantidad, "##0.00"), 7) & " " & Right(String(3, " ") & "UN", 3) & " " & Right(String(9, " ") & Format(RsImpresion!nPrecioOficial, "##0.00"), 9) & " " & Right(String(9, " ") & Format(RsImpresion!nCantidad * RsImpresion!nPrecioOficial, "##,##0.00"), 9)
                                    sLinea2 = Mid(xproducto & String(30, " "), 1, 30) & " " & Right(String(9, " ") & Format(RsImpresion!descunitario, "##,##0.00"), 9)
                                End If
                                Printer.Print sLinea
                                Printer.Print sLinea2
                            End If
                            
                          Else
                          
                            If Not (RsImpresion!lImpProdDesc) Then
                                CANT = Len(Trim(xproducto))
                                ProductoTem = xproducto
regreso5:
                                If valor = 0 Then
                                    sLinea = Right(String(8, " ") & Format(RsImpresion!nCantidad, "#,##0.00"), 8) & " " & Mid(xproducto & String(21, " "), 1, 21)
                                    Printer.Print sLinea
                                
                                    valor = 1
                                    CANT = Len(Mid(ProductoTem, 22))
                                    ProductoTem = Mid(ProductoTem, 22)
                                Else
                                    sLinea = Mid(ProductoTem & String(16, " "), 1, 16) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(6, " ") & "      ", 6) & " " & Right(String(9, " ") & "           ", 9)
                                    'sLinea = Right(String(8, " ") & "        ", 8) & " " & Mid(ProductoTem & String(21, " "), 1, 21)
                                    CANT = Len(Mid(ProductoTem, 22))
                                    ProductoTem = Mid(ProductoTem, 22)
                                    Printer.Print sLinea
                                End If
                                
                                If CANT > 0 And ximprimeDetallado = True Then
                                    GoTo regreso5
                                End If
                                If ImpDolarDelivery Then
                                    sLinea = String(8, " ") & " " & Mid(RsImpresion!tOferta & String(20, " "), 1, 20) & " " & Right(String(10, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) / nTC, "###,##0.00"), 10)
                                    'sLinea = String(8, " ") & " " & Mid(RsImpresion!tOferta & String(20, " "), 1, 20) & " " & Right(String(10, " ") & " ", 10)
                                Else
                                    sLinea = String(8, " ") & " " & Mid(RsImpresion!tOferta & String(20, " "), 1, 20) & " " & Right(String(10, " ") & Format(RsImpresion!nCantidad * RsImpresion!nPrecioOficial, "###,##0.00"), 10)
                                    'sLinea = String(8, " ") & " " & Mid(RsImpresion!tOferta & String(20, " "), 1, 20) & " " & Right(String(10, " ") & " ", 10)
                                End If
                             Printer.Print sLinea
                            End If
                            If (RsImpresion!lImpProdDesc) Then ' Para impresion con formato 2
                                If ImpDolarDelivery Then
                                    sLinea = Mid(RsImpresion!tCodigoProducto & String(8, " "), 1, 8) & " " & Right(String(7, " ") & Format(RsImpresion!nCantidad, "##0.00"), 7) & " " & Right(String(3, " ") & "UN", 3) & " " & Right(String(9, " ") & Format(RsImpresion!nPrecioOficial / nTC, "##0.00"), 9) & " " & Right(String(9, " ") & Format((RsImpresion!nCantidad * RsImpresion!nPrecioOficial) / nTC, "##,##0.00"), 9)
                                    sLinea2 = Mid(xproducto & String(30, " "), 1, 30) & " " & Right(String(9, " ") & Format((RsImpresion!descunitario) / nTC, "##,##0.00"), 9)
                                Else
                                    sLinea = Mid(RsImpresion!tCodigoProducto & String(8, " "), 1, 8) & " " & Right(String(7, " ") & Format(RsImpresion!nCantidad, "##0.00"), 7) & " " & Right(String(3, " ") & "UN", 3) & " " & Right(String(9, " ") & Format(RsImpresion!nPrecioOficial, "##0.00"), 9) & " " & Right(String(9, " ") & Format(RsImpresion!nCantidad * RsImpresion!nPrecioOficial, "##,##0.00"), 9)
                                    sLinea2 = Mid(xproducto & String(30, " "), 1, 30) & " " & Right(String(9, " ") & Format(RsImpresion!descunitario, "##,##0.00"), 9)
                                End If
                                Printer.Print sLinea
                                Printer.Print sLinea2
                            End If
                          End If
                          valor = 0
                          xTotalS = xTotalS + (RsImpresion!nCantidad * RsImpresion!nPrecioOficial)
                    End If
                       'Combos
                       If lComboDocumento And IIf(IsNull(RsImpresion!Combo), 0, Len(Trim(RsImpresion!Combo))) > 0 Then
                          If (lMostrarPrecioItemsCombo) Then
                            Isql = "select distinct nAumento from tCombo where tCombo= '" & RsImpresion!tCodigoProducto & "' and tCodigoProducto in (select tProductoCombo  from CPEDIDO where tCodigoPedido = '" & RsImpresion!tCodigoPedido & "' and tItem = '" & RsImpresion!tItem & "' and tItemCombo = '" & RsImpresion!tItemCombo & "' ) "
                            Set xRsObtAumentaPrecio = Lib.OpenRecordset(Isql, Cn)
                        
                            If (lNoVisualizarItemCombos) Then
                                If xRsObtAumentaPrecio!nAumento > 0 Then
                                    'sLinea = Mid("" & String(2, " "), 1, 2) & " " & Right(String(4, " ") & Format(RsImpresion!nCombo, "##0.00"), 4) & " " & RsImpresion!Combo & " " & Right(String(6, " ") & "(" & Format(xRsObtAumentaPrecio!nAumento * RsImpresion!nCombo, "##,##0.00") & ")", 6)
                                    sLinea = Mid("" & String(4, " "), 1, 4) & " " & Right(String(5, " ") & Format(RsImpresion!nCombo, "##0.00"), 5) & " " & Mid(RsImpresion!Combo & String(18, " "), 1, 18) & " " & Right(String(10, " ") & Format(xRsObtAumentaPrecio!nAumento * RsImpresion!nCombo, "##,##0.00"), 10)
                                    Printer.Print sLinea
                                End If
                            Else
                                'sLinea = Mid("" & String(2, " "), 1, 2) & " " & Right(String(4, " ") & Format(RsImpresion!nCombo, "##0.00"), 4) & " " & RsImpresion!Combo & " " & Right(String(6, " ") & "(" & Format(xRsObtAumentaPrecio!nAumento * RsImpresion!nCombo, "##,##0.00") & ")", 6)
                                sLinea = Mid("" & String(4, " "), 1, 4) & " " & Right(String(5, " ") & Format(RsImpresion!nCombo, "##0.00"), 5) & " " & Mid(RsImpresion!Combo & String(18, " "), 1, 18) & " " & Right(String(10, " ") & Format(xRsObtAumentaPrecio!nAumento * RsImpresion!nCombo, "##,##0.00"), 10)
                                Printer.Print sLinea
                            End If
                        Else
                            sLinea = Mid("       " & Right(String(5, " ") & Format(RsImpresion!nCombo, "#0.00"), 5) & " " & RsImpresion!Combo, 1, 35)
                            Printer.Print sLinea
                        End If
                          'Combos Propiedades
                          If lPropiedadDocumento And Len(Trim(IIf(IsNull(RsImpresion!Combo), "", RsImpresion!Combo))) > 0 Then
                          
                             xRsPropiedadCombo.Filter = "tCodigoPedido='" & RsImpresion!tCodigoPedido & "' and tItem='" & RsImpresion!tItem & "' and tItemCombo='" & RsImpresion!tItemCombo & "'"
                             If Not xRsPropiedadCombo.EOF Then
                                xRsPropiedadCombo.MoveFirst
                                Do While Not xRsPropiedadCombo.EOF
                                   'Printer.Print "        * " & xRsPropiedadCombo!Operador & " " & xRsPropiedadCombo!Descripcion
                                    ImprimeXLinea "             * " & xRsPropiedadCombo!Operador & " " & xRsPropiedadCombo!descripcion, 31, 13
                                   xRsPropiedadCombo.MoveNext
                                Loop
                             End If
                          
                             'Impresion de la Observacion
                              If lObservacionDocumento And Len(Trim(RsImpresion!tObservacionCombo)) > 0 Then
                                 ImprimeXLinea "        - " & Trim(RsImpresion!tObservacionCombo), 33, 10
                              End If
                          End If
                       End If
                       xObservacion = IIf(IsNull(RsImpresion!tObservacionpedido), "", RsImpresion!tObservacionpedido)
                       Item = RsImpresion!tCodigoPedido + RsImpresion!tItem
                       RsImpresion.MoveNext
                Loop
                    
                    If lPropiedadDocumento Then
                       xRsPropiedad.Filter = "tCodigoPedido='" & Mid(Item, 1, 10) & "' and tItem='" & Mid(Item, 11, 3) & "'"
                       If Not xRsPropiedad.EOF Then
                          xRsPropiedad.MoveFirst
                          Do While Not xRsPropiedad.EOF
                              ImprimeXLinea "        * " & IIf(xRsPropiedad!nCantidad = 1, "", "(" & xRsPropiedad!nCantidad & ")") & xRsPropiedad!Operador & " " & xRsPropiedad!descripcion, 31, 10
                             xRsPropiedad.MoveNext
                          Loop
                       End If
                                   
                       'Impresion de la Observacion
                       If lObservacionDocumento And Len(Trim(xObservacion)) > 0 Then
                          ImprimeXLinea "        - " & Trim(xObservacion), 33, 2
                       End If
                    End If

                    Printer.Print String(40, "-")
                    RsImpresion.MoveLast
                    Dim xGravada As Double
                    Dim xInafecta As Double
                    If ImpDolarDelivery Then
                    
                        Printer.Print String(40, " ")
                        xGravada = Calcular("SELECT SUM(D.nPrecioNeto*nCantidad) As Codigo FROM dDocumento D WHERE D.tDocumento = '" & RsImpresion!tDocumento & "' AND round(D.nPrecioImpuesto1,4)<> 0 ", Cn)
                        xInafecta = Calcular("select  [dbo].[ObteniendoExoIna] (2, '" & RsImpresion!tDocumento & "') Codigo", Cn)
                        xExonerado = Calcular("select  [dbo].[ObteniendoExoIna] (1, '" & RsImpresion!tDocumento & "') Codigo", Cn)

                        If Not (IsNull(RsImpresion!lOpGravInaf) Or RsImpresion!lOpGravInaf = 0) Then
                            Printer.Print Right(String(27, " ") & "Neto Gravado " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, xGravada / nTC), "##,###,##0.00"), 13)
                        End If
                        
                        If RsImpresion!nRecargo > 0 Then
                           'Printer.Print
                           Printer.Print Right(String(27, " ") & "Recargo " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nRecargo / nTC), "##,###,##0.00"), 13)
                        End If
                    Else
                        'FE SPRING

                        Printer.Print String(40, " ")
    '
                        xGravada = Calcular("SELECT SUM(D.nPrecioNeto*nCantidad) As Codigo FROM dDocumento D WHERE D.tDocumento = '" & RsImpresion!tDocumento & "' AND round(D.nPrecioImpuesto1,4)<> 0 ", Cn)
                        xInafecta = Calcular("select  [dbo].[ObteniendoExoIna] (2, '" & RsImpresion!tDocumento & "') Codigo", Cn)
                        xExonerado = Calcular("select  [dbo].[ObteniendoExoIna] (1, '" & RsImpresion!tDocumento & "') Codigo", Cn)
                        
                        If Not (IsNull(RsImpresion!lOpGravInaf) Or RsImpresion!lOpGravInaf = 0) Then
                            Printer.Print Right(String(27, " ") & "Neto Gravado " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, xGravada), "##,###,##0.00"), 13)
                        End If
                        
                        If RsImpresion!nRecargo > 0 Then
                           'Printer.Print
                           Printer.Print Right(String(27, " ") & "Recargo " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nRecargo), "##,###,##0.00"), 13)
                        End If

                    End If

                    nSumaNeto = RsImpresion!nNeto
                    
                    If IsNull(RsImpresion!lImpuesto1) Or Not RsImpresion!lImpuesto1 Then
                       nSumaNeto = nSumaNeto + RsImpresion!nprecioImpuesto1
                    End If
                    
                    If IsNull(RsImpresion!lImpuesto2) Or Not RsImpresion!lImpuesto2 Then
                       nSumaNeto = nSumaNeto + RsImpresion!nprecioImpuesto2
                    End If
                    
                    If IsNull(RsImpresion!lImpuesto3) Or Not RsImpresion!lImpuesto3 Then
                       nSumaNeto = nSumaNeto + RsImpresion!nprecioImpuesto3
                    End If
                                                  
                    Dim xValorLey1  As Double
                    Dim xCantidadLey1  As Double
                    xValorLey1 = Calcular("SELECT ISNULL(SUM(D.nLey1),0) As Codigo FROM VDocumento D WHERE  D.tDocumento = '" & RsImpresion!tDocumento & "'", Cn)
                    xCantidadLey1 = Calcular("SELECT ISNULL(SUM(D.ncantidad),0) As Codigo FROM VDocumento D WHERE D.lLey1=1 and D.tDocumento = '" & RsImpresion!tDocumento & "'", Cn)
                                                  
                    If ImpDolarDelivery Then
                        If Not (IsNull(RsImpresion!lImpuesto1) Or RsImpresion!lImpuesto1 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto1 & " " & Format(nPorcentaje1, "##0.00") & "% " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto1 / nTC), "##,###,##0.00"), 13)
                        End If
                           
                        If Not (IsNull(RsImpresion!lImpuesto2) Or RsImpresion!lImpuesto2 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto2 & " " & Format(nPorcentaje2, "##0.00") & "% " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto2 / nTC), "##,###,##0.00"), 13)
                        End If
                        
                        If Not (IsNull(RsImpresion!lImpuesto3) Or RsImpresion!lImpuesto3 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto3 & " " & Format(nPorcentaje3, "##0.00") & "% " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto3 / nTC), "##,###,##0.00"), 13)
                        End If
                        
                        If sLey1 <> "" Then
                            Printer.Print Right(String(27, " ") & sLey1 & "  " & sMonE & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, xValorLey1 / nTC), "##,###,##0.00"), 13)
                        End If
                        
                         If lPagoAntesImpresion Then
                            lcuentaCobrar = True
    
                            If lcuentaCobrar And lPagocortesiaAI = False Then
                              Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta / nTC, "##,###,##0.00"), 13)
                            Else
                              Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(0, "##,###,##0.00"), 13)
                            End If
                        Else
                        
                          Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta / nTC, "##,###,##0.00"), 13)
                        End If
                    Else
                        If Not (IsNull(RsImpresion!lImpuesto1) Or RsImpresion!lImpuesto1 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto1 & " " & Format(nPorcentaje1, "##0.00") & "% " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto1), "##,###,##0.00"), 13)
                        End If
                           
                        If Not (IsNull(RsImpresion!lImpuesto2) Or RsImpresion!lImpuesto2 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto2 & " " & Format(nPorcentaje2, "##0.00") & "% " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto2), "##,###,##0.00"), 13)
                        End If
                        
                        If Not (IsNull(RsImpresion!lImpuesto3) Or RsImpresion!lImpuesto3 = 0) Then
                           Printer.Print Right(String(27, " ") & sImpuesto3 & " " & Format(nPorcentaje3, "##0.00") & "% " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, RsImpresion!nprecioImpuesto3), "##,###,##0.00"), 13)
                        End If
                        
                        If sLey1 <> "" Then
                            Printer.Print Right(String(27, " ") & sLey1 & "  " & sMonN & " : ", 27) & Right(String(13, " ") & Format(IIf(lPagocortesiaAI And lPagoAntesImpresion, 0, xValorLey1), "##,###,##0.00"), 13)
                        End If
                        
                         If lPagoAntesImpresion Then
                            lcuentaCobrar = True
    
                            If lcuentaCobrar And lPagocortesiaAI = False Then
                              Printer.Print Right(String(27, " ") & "Importe Total " & sMonN & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta, "##,###,##0.00"), 13)
                            Else
                              Printer.Print Right(String(27, " ") & "Importe Total " & sMonN & " : ", 27) & Right(String(13, " ") & Format(0, "##,###,##0.00"), 13)
                            End If
                        Else
                        
                          Printer.Print Right(String(27, " ") & "Importe Total " & sMonN & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta, "##,###,##0.00"), 13)
                        End If

                        If lEquivalencia = True Or lDocumEquivaPrecuenta = True Then
                            If lPagoAntesImpresion Then
                                lcuentaCobrar = True
        
                                If lcuentaCobrar And lPagocortesiaAI = False Then
                                    Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta / nTC, "##,###,##0.00"), 13)
                                    Printer.Print
                                Else
                                    Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(0, "##,###,##0.00"), 13)
                                    Printer.Print
                                End If
                            Else
                                    
                                Printer.Print Right(String(27, " ") & "Importe Total " & sMonE & " : ", 27) & Right(String(13, " ") & Format(RsImpresion!nVenta / nTC, "##,###,##0.00"), 13)
                                Printer.Print
                            End If
                        End If
                    End If
                    'TEXTO AGENTE RETENCION
'                    Printer.Print " "
'                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & RsImpresion!tDocumento & "')", Cn) = 1 Then
'                        ImprimeXLinea tTextoAgenteRetencion, 40, 5
'                    End If
'                    Printer.Print " "
'                    If ImpDolarDelivery Then
'                        If lPagoAntesImpresion Then
'                                lcuentaCobrar = True
'
'                            If lcuentaCobrar And lPagocortesiaAI = False Then
'                                txt = "Son: " & NumeroCadena(str(RsImpresion!nVenta / nTC)) & " " & sMonedaE
'                                ImprimeXLinea txt, 40, 5
'
'                            Else
'                                txt = "Son: " & NumeroCadena(str(0)) & " " & sMonedaE
'                                ImprimeXLinea txt, 40, 5
'                            End If
'                        Else
'                            txt = "Son: " & NumeroCadena(str(RsImpresion!nVenta / nTC)) & " " & sMonedaE
'                            ImprimeXLinea txt, 40, 5
'                        End If
'                      Else
'                        If lPagoAntesImpresion Then
'                                lcuentaCobrar = True
'
'                            If lcuentaCobrar And lPagocortesiaAI = False Then
'                                txt = "Son: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
'                                ImprimeXLinea txt, 40, 5
'
'                            Else
'                                txt = "Son: " & NumeroCadena(str(0)) & " " & sMonedaN
'                                ImprimeXLinea txt, 40, 5
'                            End If
'                        Else
'                            txt = "Son: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
'                            ImprimeXLinea txt, 40, 5
'                        End If
'                      End If
    '-------------------------------------GCAA 29/02/2020 ---- TANTA ---- FACTURA-------------------
                    RsImpresion.MoveFirst
'                    Dim CantidaBonificacion As Integer
'                    CantidaBonificacion = 0
'                    Dim xacumulado As Double
'                    xacumulado = 0
'                    Do While Not RsImpresion.EOF
'
'                        If Round(RsImpresion!nPrecioOficial, 2) = 0 Then
'                            CantidaBonificacion = CantidaBonificacion + 1
'                            If CantidaBonificacion = 1 Then
'                                Printer.Print String(40, "-")
'                                Printer.Print "BONIFICACIÓN"
'                            End If
'
'                            Dim NombreProducto As String
'                            If ximprimeDetallado = False Then
'                                NombreProducto = RsImpresion!Producto
'                            Else
'                                NombreProducto = RsImpresion!ProductoDetallado
'                            End If
'
'                            sLinea = Mid(NombreProducto & String(14, " "), 1, 14) & " " & Right(String(6, " ") & Format(RsImpresion!nCantidad, "##0.00"), 6) & " " & Right(String(7, " ") & Format(RsImpresion!nPrecioVentaRefer, "###0.00"), 7) & " " & Right(String(10, " ") & Format(RsImpresion!nCantidad * RsImpresion!nPrecioVentaRefer, "##,##0.00"), 10)
'                            Printer.Print sLinea
'                              xacumulado = xacumulado + (RsImpresion!nCantidad * RsImpresion!nPrecioVentaRefer)
'                        End If
'
'                        RsImpresion.MoveNext
'                    Loop
                   
'                    If CantidaBonificacion > 0 Then
'                        Printer.Print Right(String(27, " ") & "Total de Bonificación   ", 27)
'                        Printer.Print Right(String(27, " ") & "         sin impuestos: ", 27) & Right(String(13, " ") & ((Format(xacumulado, "##,##0.00") - (Format(xacumulado, "##,##0.00") * nPorcentaje1) / 100)), 13)
'                        Printer.Print Right(String(27, " ") & "Total de Bonificación : ", 27) & Right(String(13, " ") & Format(xacumulado, "##,##0.00"), 13)
'
'                        Printer.Print String(40, "-")
'                        Printer.Print ""
'                    End If
                    RsImpresion.MoveLast
             
                   If lPagoAntesImpresion And RsImpresion!nVenta > 0 Then
                        lcuentaCobrar = True
                        Set RsPagoDocumento = Lib.OpenRecordset("exec usp_Inforest_Impresion '" & RsImpresion!tDocumento & "', 9 ", Cn)
                        Printer.Print " "
                            If RsPagoDocumento.RecordCount > 0 Then
'                                If (lImprimirFormaPago) Then
'                                Printer.Print String(40, "-")
'                                Printer.Print "Tipo de Pago: AL CONTADO"
'                                Printer.Print String(40, "-")
                                'End If
                                For i = 0 To RsPagoDocumento.RecordCount - 1
                                    lcuentaCobrar = False
                                        Printer.Print "Tipo de Pago: " & RsPagoDocumento!pago & " " & Format(RsPagoDocumento!Monto, "##,###,##0.00")
                                    'GCAA 09032021
                                    RsPagoDocumento.MoveNext
                                Next i
                                RsPagoDocumento.MoveFirst
                                
                                Printer.Print "Vuelto       :" & sMonN & " " & Format(RsPagoDocumento!Vuelto, "##,###,##0.00")
                            Else
                                '''''''''''''''''
'                                If (lImprimirFormaPago) Then
'                                    Printer.Print String(40, "-")
'
'                                    If Calcular(" select isnull(lCanalDelivery,0) as  Codigo from TCANALVENTA where tCodigoCanalVenta ='" & RsImpresion!tTipoPedido & "'", Cn) Then
'                                        Printer.Print "Forma de Pago: AL CONTADO"
'                                    Else
'                                        Printer.Print "Forma de Pago: AL CREDITO"
'                                    End If
'                                    Printer.Print String(40, "-")
'                                End If
                                Printer.Print ""
                                If lcuentaCobrar And lPagocortesiaAI = False Then
                                    Printer.Print "Tipo de Pago: Enviado a Cuenta por Cobrar"
                                Else
                                    ImprimeXLinea "Tipo de Pago: " & lGlosaTrans, 40, 15
                                End If
                            End If
                        Printer.Print " "
                    Else
                        If (lImprimirFormaPago) Then
                            Printer.Print String(40, "-")
                            'Printer.Print "Tipo de Pago: AL CONTADO"
                            Printer.Print String(40, "-")
                           
                        End If
                    End If
                    If Not lImpClienteCabDoc Then
                        Printer.Print String(40, "-")
'                        If Calcular("select upper(tResumido) as Codigo from vTipoIdentidad inner join TCLIENTE on vTipoIdentidad.Codigo=tcliente.tTipoIdentidad where tcliente.tCodigoCliente= (select isnull(tCodigoCliente,'') as codigo From MDOCUMENTO where tDocumento='" & RsImpresion!tDocumento & "')", Cn) = "DNI" Then
'                            ImprimeXLinea "Razon Social: " & RsImpresion!Cliente, 40, 9
'                            Printer.Print "CUIT        : " & RsImpresion!Ruc
'                            ImprimeXLinea "Direccion   : " & RsImpresion!direccion, 40, 9
'                        Else
'                            ImprimeXLinea "Empresa: " & RsImpresion!Cliente, 40, 9
'                            Printer.Print "RUC    : " & RsImpresion!Ruc
'                            ImprimeXLinea "Direcc.: " & RsImpresion!direccion, 40, 9
'                        End If
                        Dim xTipoIdentidad As String
                         Printer.Print " "
                        xTipoIdentidad = ""
                        xTipoIdentidad = Calcular("select upper(tResumido) as Codigo from vTipoIdentidad inner join TCLIENTE on vTipoIdentidad.Codigo=tcliente.tTipoIdentidad where tcliente.tCodigoCliente= (select isnull(tCodigoCliente,'') as codigo From MDOCUMENTO where tDocumento='" & RsImpresion!tDocumento & "')", Cn)
                        If xTipoIdentidad = "0" Then
                            xTipoIdentidad = "Iden."
                        End If
                        Printer.Print xTipoIdentidad & "    : " & RsImpresion!Ruc
                        ImprimeXLinea "Razon Social: " & RsImpresion!Cliente, 40, 9
                        ImprimeXLinea "Direccion   : " & RsImpresion!direccion, 40, 9
                         Printer.Print " "
                        
                        
                        
                        Printer.Print String(40, "-")
                    End If
                        Printer.Print " "
                        'Printer.Print " "
                        'Printer.Print " "
                    If CodigoHash <> 0 Then
                        'ELDCQ 02/05/2022
                        ImprimeXCentro "Comprobante Autorizado", 40
                        Printer.Print " "
                        xImpresionFE = Calcular(" SELECT isnull(tImpresionFE,'') as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & RsImpresion!tDocumento & "')", Cn)
                        If xImpresionFE <> "" Then
                            ImprimeXLinea xImpresionFE, 40, 0
                            Printer.Print " "
                        End If
                        Printer.PaintPicture CodigoHash, 900, Printer.CurrentY, 2000, 2000
                        Dim d As Integer
                        For d = 1 To 10
                            Printer.Print
                        Next d
                        Printer.Print " "
                        Dim CAE As String
                        CAE = Calcular(" SELECT top 1 isnull(CAE,'') as Codigo from  DOCUMENTO_ARGENTINA WHERE Documento='" & RsImpresion!tDocumento & "'", Cn)
                        Dim FVencimientoCAE As String
                        FVencimientoCAE = Format(Calcular(" SELECT top 1 isnull(FechaVenceCAE,GETDATE()) as Codigo from  DOCUMENTO_ARGENTINA WHERE Documento='" & RsImpresion!tDocumento & "'", Cn), "dd/MM/yyyy")
                        
                        ImprimeXLinea "CAE Nº: " & CAE, 40, 8
                        ImprimeXLinea "Fecha Vencimiento CAE: " & FVencimientoCAE, 40, 0
                        Printer.Print " "
                        'Fin ELDCQ 02/05/2022
                     Else
                        If lImpresionCodigoBarras Then
                           Printer.PaintPicture CodigoHash, 250, Printer.CurrentY, 3400, 1200
                           Printer.CurrentY = Printer.CurrentY + 1200
                        Else
                           ImprimeXCentro cadenaCodigoHash, 40
                        End If
                     End If
            
                     ImprimeXCentro sPie, 40
                     
                     
                     If lImprimeImagPieDocumento Then
                            If pie <> 0 Then
                                Printer.Print " "
                                Printer.PaintPicture pie, 500, Printer.CurrentY, 2860, 1600
                                Dim i1 As Integer
                                For i1 = 1 To 16
                                    Printer.Print
                                Next i1
                            End If
                    End If
                        
                    Printer.Print
                    Printer.EndDoc
        End Select
End Sub




Private Function TicketFacturaFE(IFEpson As PrinterFiscal, RsDatos As Recordset) As Integer
    
    Dim respuesta As Boolean
    Dim RespVALOR As Integer
    respuesta = True
    RespVALOR = 0
    
On Error GoTo fin

    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600

    Dim vTipoResponsable As String
    Dim TDComprador As String
    Dim LetraDoc As String
    LetraDoc = Left(RsDatos!tDocumento, 1)
    vTipoResponsable = Calcular("select top 1 isnull(vTR,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn)
    TDComprador = Left(Calcular("select top 1 isnull(TipoIdentidad,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn), 6)
    
    If respuesta Then respuesta = IFEpson.OpenInvoice("T", "C", LetraDoc, "1", "P", "12", "I", vTipoResponsable, Left(RsDatos!Cliente, 40), IIf(Left(Trim(RsDatos!direccion), 40) = "", "-", Left(Trim(RsDatos!direccion), 40)), TDComprador, Replace(RsDatos!Ruc, "-", ""), "N", String(40, "-"), "147 Telefono gratuito CABA, Area de ", "Defensa y Proteccion al Consumidor", String(40, "-"), "Obs. " & Left(IIf(RsDatos!tObservacion = "", Right(RsDatos!tDocumento, 8), RsDatos!tObservacion), 35), "G")
    
    If respuesta Then RespVALOR = 1
    
    Do While Not RsDatos.EOF
        If respuesta Then respuesta = IFEpson.SendInvoiceItem(Left(RsDatos!Producto, 20), Replace(Format(RsDatos!nCantidad, "0####.##0"), ".", ""), IIf(LetraDoc = "B", Replace(Format(RsDatos!nPrecioOficial, "0######.#0"), ".", ""), Replace(Format((RsDatos!nprecioVenta / (1 + (nPorcentaje1 + nPorcentaje2 + nPorcentaje3) / 100) + RsDatos!descunitario / RsDatos!nCantidad), "0######.#0"), ".", "")), Replace(Format(nPorcentaje1 * 100, "0##0"), ".", ""), "M", "0", "0", "", "", "", "1050", "0")
     RsDatos.MoveNext
    Loop
    RsDatos.MoveLast

    If respuesta Then respuesta = IFEpson.GetInvoiceSubtotal("P", "SUB TOTAL")
    
    If respuesta Then RespVALOR = 2
    
    If RsDatos!MotivoDescuento <> "" Then
        'ImprimeXLinea "Descuento     : " & RsImpresion!MotivoDescuento, 40, 16
        If RsDatos!nDescuento > 0 Then
            If respuesta Then respuesta = IFEpson.SendInvoicePayment(Left(RsDatos!MotivoDescuento, 22), Replace(Format(RsDatos!nDescuento, "0#########.#0"), ".", ""), "D")
        End If
    Else
        If RsDatos!nDescuento > 0 Then
            If respuesta Then respuesta = IFEpson.SendInvoicePayment("Descuento: ", Replace(Format(RsDatos!nDescuento, "0#########.#0"), ".", ""), "D")
        End If
    End If

    If lPagoAntesImpresion And RsDatos!nVenta > 0 Then
    Dim RsPagoDocumento As Recordset
    Dim i As Integer
    Set RsPagoDocumento = Lib.OpenRecordset("exec usp_Inforest_Impresion '" & RsDatos!tDocumento & "', 9 ", Cn)
        If RsPagoDocumento.RecordCount > 0 Then
            For i = 0 To RsPagoDocumento.RecordCount - 1
                If respuesta Then respuesta = IFEpson.SendInvoicePayment(RsPagoDocumento!pago + "(" + Trim(CStr(RsPagoDocumento!ValorReal)) + ")", Replace(Format(RsPagoDocumento!MN, "0#########.#0"), ".", ""), "T")
                RsPagoDocumento.MoveNext
            Next i
        End If
    End If
    
    If respuesta Then respuesta = IFEpson.CloseInvoice("T", LetraDoc, "FINALIZA")
    
    
    
    'If Respuesta Then Respuesta = IFEpson.CutPaper()
    TicketFacturaFE = RespVALOR
    Exit Function
fin:
    MsgBox "Mensaje: " & error, vbInformation, "Facturacion impresora fiscal"
    TicketFacturaFE = RespVALOR
End Function
Private Function TicketFacturaNC(IFEpson As PrinterFiscal, RsDatos As Recordset) As Integer
    Dim respuesta As Boolean
    Dim RespVALOR As Integer
    respuesta = True
    RespVALOR = 0

On Error GoTo fin
    Dim CajaIF As String
    CajaIF = Calcular("select isnull(timpfiscal,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
    IFEpson.PortNumber = Calcular("select isnull(tPuertoSerial,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn)
    IFEpson.BaudRate = Calcular("select isnull(tBitPorSegundo,'') as Codigo from timpresora where timpresora='" & CajaIF & "' and tcaja='" & sCaja & "'", Cn) '9600

    Dim vTipoResponsable As String
    Dim TDComprador As String
    Dim LetraDoc As String
    LetraDoc = Left(RsDatos!tDocumento, 1)
    vTipoResponsable = Calcular("select top 1 isnull(vTR,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn)
    TDComprador = Left(Calcular("select top 1 isnull(TipoIdentidad,'') as codigo from vCliente where tidentidad='" & RsDatos!Ruc & "'", Cn), 6) '
    
    If respuesta Then respuesta = IFEpson.OpenInvoice("M", "C", LetraDoc, "1", "P", "12", "I", vTipoResponsable, Left(RsDatos!Cliente, 40), Left(Trim(RsDatos!direccion), 40), TDComprador, Replace(RsDatos!Ruc, "-", ""), "N", String(40, "-"), "Documento de Referencia:", Left(RsDatos!tDocumento, 1) & Right(Left(RsDatos!tDocumento, 6), 4) & "-" & Right(RsDatos!tDocumento, 8), String(40, "-"), Left("Motivo: " & RsDatos!tObservacion, 40), "G")
    
    If respuesta Then RespVALOR = 1
    
    Do While Not RsDatos.EOF
        If respuesta Then respuesta = IFEpson.SendInvoiceItem(Left(RsDatos!Producto, 20), Replace(Format(RsDatos!nCantidad, "0####.##0"), ".", ""), IIf(LetraDoc = "B", Replace(Format(RsDatos!nprecioVenta, "0######.#0"), ".", ""), Replace(Format(RsDatos!nPrecioNeto, "0######.#0"), ".", "")), Replace(Format(nPorcentaje1 * 100, "0##0"), ".", ""), "M", "0", "0", "", "", "", "1050", "0")
     RsDatos.MoveNext
    Loop
    RsDatos.MoveLast
    If respuesta Then respuesta = IFEpson.GetInvoiceSubtotal("P", "SUB TOTAL")
    
    If respuesta Then RespVALOR = 2
    
    If lPagoAntesImpresion And RsDatos!nVenta > 0 Then
    Dim RsPagoDocumento As Recordset
    Dim i As Integer
    Set RsPagoDocumento = Lib.OpenRecordset("exec usp_Inforest_Impresion '" & RsDatos!tDocumento & "', 9 ", Cn)
        If RsPagoDocumento.RecordCount > 0 Then
            For i = 0 To RsPagoDocumento.RecordCount - 1
                If respuesta Then respuesta = IFEpson.SendInvoicePayment(RsPagoDocumento!pago + "(" + Trim(CStr(RsPagoDocumento!ValorReal)) + ")", Replace(Format(RsPagoDocumento!MN, "0#########.#0"), ".", ""), "T")
                RsPagoDocumento.MoveNext
            Next i
        End If
    End If
    
    If respuesta Then respuesta = IFEpson.CloseInvoice("M", LetraDoc, "FINALIZA")
    
    
    TicketFacturaNC = RespVALOR
    Exit Function
fin:
    MsgBox "Mensaje: " & error, vbInformation, "Facturacion impresora fiscal"
    TicketFacturaNC = RespVALOR
End Function

