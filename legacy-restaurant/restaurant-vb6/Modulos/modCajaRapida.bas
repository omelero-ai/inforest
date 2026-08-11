Attribute VB_Name = "modCajaRapida"
Public nSalir As Long
Public hk As New License
Sub Main()
 Dim existe As Boolean
 Dim AnoMesSv As String
 Dim directorioSistema  As String
 Dim archI As String
 
      
    'dic 2010 un exe varias bd
    ultimoConectado = True
    moduloUso = "PuntoVenta"
    'dic 2010
    
  '====================================== CENTRALIZADA
  Dim verificaAdministracionCentralizada As String
    
    verificaAdministracionCentralizada = Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "CENTRALIZADA", "OFF"))
    If verificaAdministracionCentralizada = "ON" Then ' si es "ON" significa que jala informacion del servidor central
        lCentral = True  '  flag de adm. centralizada
        sServidorCentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "SERVIDOR", "0.0.0.0")) ' leer servidor central
        bdInforestCentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "BASEDATO", "INFOREST")) '  leer base de datos en servidor central, a consultar para actualizar datos
    End If
    
    '========================================== INFOREST CENTRAL GCAA 26/02/2020
    
     VerConexionInfoRestCentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "InfoRestCentral", "CENTRAL", "OFF"))
    If VerConexionInfoRestCentral = "ON" Then ' si es "ON" significa que jala informacion del servidor central inforest
        xlCentral = True  '  flag de adm. centralizada
        xservidorcentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "InfoRestCentral", "SERVIDOR", "0.0.0.0")) ' leer servidor central
        xbasecenntral = Trim(LeerIni(App.Path + "\INFOREST.INI", "InfoRestCentral", "BASEDATO", "INFORESTCD")) 'leer base de datos en servidor central, a consultar para actualizar datos
    Else
        xlCentral = False  '  flag de adm. centralizada
        xservidorcentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "InfoRestCentral", "SERVIDOR", "0.0.0.0")) ' leer servidor central
        xbasecenntral = Trim(LeerIni(App.Path + "\INFOREST.INI", "InfoRestCentral", "BASEDATO", "INFORESTCD")) 'leer base de datos en servidor central, a consultar para actualizar datos
    End If
    
    
    '--------------- gcaa 06012021
'     If App.PrevInstance Then
'          CloseWindowsByClassName ("InfoRest.exe")
'     End If
    
    If App.PrevInstance Then
        MsgBox "Ya se esta ejecutando el Aplicativo!", vbInformation, "Atención!!!"
        End
    End If
    '---------------
   lVersionEducativa = False
    
   lHARDkey = False
        
  'Configuracion ini
   sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
   sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
   sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
   sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
   sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
  
   sRutaCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "SERVIDOR", "."))
   sMDBCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "BASEDATO", "INFOREST"))
   
   sRutaWAP = Trim(LeerIni(App.Path + "\INFOREST.INI", "WebMobile", "SERVIDOR", "."))
   sMDBWAP = Trim(LeerIni(App.Path + "\INFOREST.INI", "WebMobile", "BASEDATO", "WEBAPP"))
   
   MensajeVisaQR = Trim(LeerIni(App.Path + "\INFOREST.INI", "VisaQR", "Mensaje", "Ahora escanea el codigo QR y paga tu recibo con tu billetera electronica Yape, Tunki, Lukita, MobileCard o Ligo."))
   'FACTURACION_E_PERU
   sRutaFE = Trim(LeerIni(App.Path + "\FACTURACION.INI", "Conexion", "SERVIDOR", "."))
   sMDBFE = Trim(LeerIni(App.Path + "\FACTURACION.INI", "Conexion", "BASEDATO", "BDEFACT"))
   '-----------------
   
   TipoFacturacion = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "TIPOFACTURACION", "0"))
   
   AvisoLicencia = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "AVISO", 1))

   Dim DContin As String
   DContin = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "DIASCONTINGENCIA", "7"))
   
   RutaConsultaRuc = Trim(LeerIni(App.Path + "\INFOREST.INI", "RUTAS", "RutaRuc", "http://services.wijoata.com/consultar-ruc/api/ruc"))
   
  ' RutaConsultaRuc = Trim(LeerIni(App.Path + "\INFOREST.INI", "RUTAS", "RutaRuc", "http://161.132.106.116/WsConsultaRuc/Infomatica.svc/GetClienteRuc/3"))
   VisibleTC = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "VISIBLETC", 1))
   
   VisibleTC = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "VISIBLETC", 1))
   
   If IsNumeric(DContin) Then
    DiasContingenciaDocumentos = CDbl(DContin)
   Else
    DiasContingenciaDocumentos = 7
   End If
   
   
   Dim RsParametro As Recordset
   Dim RsCaja As Recordset
   Dim RsBusca As Recordset
   Dim RsFactura As Recordset
   Dim RsFpape As Recordset
   Set Lib = New Libreria16.Applications
   
   'Configuración
   sUserName = "infhotel"
   sUserPassword = "4gust1n-fl0r14n"

    'auditoria
    nCorrelativoAcceso = 0
    tModuloSeg = "01" 'XX= CODIGO DEL MODULO DE LA TABLA DE MMODULO DE SEGURIDAD
    Set CnSeg = New Connection
      
    CnSeg.Provider = "SQLOLEDB"
    CnSeg.CursorLocation = adUseServer
    CnSeg.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=INFSEGURIDAD"
    CnSeg.CommandTimeout = 250
    
    
    If xlCentral = True Then
         'inforest central GCAA 26/02/2020
        Set Conex = New Connection
        Conex.Provider = "SQLOLEDB"
        Conex.CursorLocation = adUseServer
        Conex.ConnectionString = "User ID=" & sUserName & _
                              ";password=" & sUserPassword & _
                              ";Data Source=" & xservidorcentral & _
                              ";Initial Catalog=" & xbasecenntral
        Conex.CommandTimeout = 250
        Conex.Open
         
         Set xcnDefault = Conex
    End If
    
    
    
    'auditoria
    
    
   Set Cn = New Connection
   Cn.Provider = "SQLOLEDB"
   Cn.CursorLocation = adUseServer
   Cn.ConnectionString = "User ID=" & sUserName & _
                         ";password=" & sUserPassword & _
                         ";Data Source=" & sRuta & _
                         ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 250
   Cn.Open
    
    Set cnDefault = Cn
   
    '====================================================
    If lVersionEducativa = True Then
     directorioSistema = ObtenerDirectorioSO
     directorioSistema = directorioSistema & "\Infhotel.csg"
     existe = FileExists(directorioSistema)
     If existe = True Then
        AnoMesSv = Calcular("select dbo.returnAnoMes(getdate()) as codigo", Cn)
        archI = obtieneAnoMes(directorioSistema)
        If val(AnoMesSv) > val(archI) Then
            MsgBox "Error Fatal: Su Licencia Académica a Caducado. Comuníquese con Personal De INFHOTEL", vbCritical, sMensaje
            End
        End If
     Else
        MsgBox "Error Fatal: Comuníquese con Personal De INFHOTEL", vbCritical, sMensaje
        End
     End If
    End If
    
    'extranjeroBolivia
    '=================================================
    pais = ObtienePais
    '=================================================
    
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "INFOREST"
   sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   lLongitud = IIf(IsNull(RsParametro!lLongitud), False, RsParametro!lLongitud)
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)

   '===================================================================
   '------VERIFICACION DE VERSION Actualización automática
        
    Dim sVersion As String
    Dim sVersionExe As String
    Dim RsVersion As Recordset
    
    Dim dteCreate As Date
    Dim dteAccess As Date
    Dim dteModify As Date
    
    Dim dteCreateACT As Date
    Dim dteAccessACT As Date
    Dim dteModifyACT As Date
    
    Dim mstrFile As String
    mstrFile = App.Path + "\" & App.EXEName & ".exe"
    
    Dim sActRuta As String
    sActRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Actualizador", "RUTA", ".")) & "\" & App.EXEName & ".exe"
    
     'If FileExists(mstrFile) Then
'         Call MsgBox("Enter a file path.", vbCritical, "No File Specified")
'         Exit Sub
'        If Not GetFileTimes(mstrFile, dteCreate, dteAccess, dteModify, True) Then
'            Call MsgBox("Error getting file dates.", vbCritical, "Unable to Read File Dates")
'            End
'        End If
    
     'End If
    'If FileExists(sActRuta) Then
'         Call MsgBox("Enter a file path.", vbCritical, "No File Specified")
'         Exit Sub
'        If Not GetFileTimes(sActRuta, dteCreateACT, dteAccessACT, dteModifyACT, True) Then
'            'Call MsgBox("Error getting file dates.", vbCritical, "Unable to Read File Dates")
'            'End
'            dteModifyACT = dteModify
'        End If
     'Else
        
     'End If

    sVersion = ""
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    sVersion = Calcular("SELECT tVersion As Codigo FROM tParametro", Cn)
    
'    If sVersion <> sVersionExe Or (sVersion = sVersionExe And dteModifyACT > dteModify) Then
'        MsgBox "Existe Una Nueva Actualizacion Disponible", vbInformation, sMensaje
'
'        'CREA LA CARPETA PARA BACKUP DE EXES
'        Dim Backup As String
'        Backup = App.Path + "\ExesHistoricos"
'
'        'CREA LA CARPETA Y VALIDA QUE NO EXISTA
'        If Dir(Backup, vbDirectory) = "" Then
'           MkDir (Backup)
'        End If
'
'        Shell App.Path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
'        End
'    End If
       
   '------
   
   'validacion acceso Licencias
    'validacionLicenciasInfhotel
   '===================================================================

   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!IMPUESTO1), 0, RsParametro!IMPUESTO1)
   nPorcentaje2 = IIf(IsNull(RsParametro!IMPUESTO2), 0, RsParametro!IMPUESTO2)
   nPorcentaje3 = IIf(IsNull(RsParametro!IMPUESTO3), 0, RsParametro!IMPUESTO3)
   nTiempo = IIf(IsNull(RsParametro!nTiempo), 0, RsParametro!nTiempo)
   nChkTiempo = IIf(IsNull(RsParametro!nChkTiempo), 0, RsParametro!nChkTiempo)
   nDELIVERY = IIf(IsNull(RsParametro!nDELIVERY), 0, RsParametro!nDELIVERY)
   nLlevar = IIf(IsNull(RsParametro!nLlevar), 0, RsParametro!nLlevar)
   nCanal4 = IIf(IsNull(RsParametro!nCanal4), 0, RsParametro!nCanal4)
   nCanal5 = IIf(IsNull(RsParametro!nCanal5), 0, RsParametro!nCanal5)
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sPie = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   
   sInfacto = IIf(IsNull(RsParametro!tGlosaImprInafecto), "", RsParametro!tGlosaImprInafecto)
   sExonerado = IIf(IsNull(RsParametro!tGlosaImprExonerado), "", RsParametro!tGlosaImprExonerado)
   
   sPiePreCuenta = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sFax = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
   sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   sMail = IIf(IsNull(RsParametro!tEmail), "", RsParametro!tEmail)
   sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   sMonE2 = IIf(IsNull(RsParametro!tMonE2), "", RsParametro!tMonE2)
   sMonE3 = IIf(IsNull(RsParametro!tMonE3), "", RsParametro!tMonE3)
   sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   sMonedaE2 = IIf(IsNull(RsParametro!tMonedaE2), "", RsParametro!tMonedaE2)
   sMonedaE3 = IIf(IsNull(RsParametro!tMonedaE3), "", RsParametro!tMonedaE3)
   sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tpassword), "", RsParametro!tpassword)))
   sElimina = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
   nFItem = IIf(IsNull(RsParametro!nItem), 0, RsParametro!nItem)
   nCabecera = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
   nDetalle = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
   nFItemGuia = IIf(IsNull(RsParametro!nItemGuia), 0, RsParametro!nItemGuia)
   nCabeceraGuia = IIf(IsNull(RsParametro!nCabeceraGuia), 0, RsParametro!nCabeceraGuia)
   nDetalleGuia = IIf(IsNull(RsParametro!nDetalleGuia), 0, RsParametro!nDetalleGuia)
   lPrinter = IIf(IsNull(RsParametro!lPrinter), False, RsParametro!lPrinter)
   lRapido = IIf(IsNull(RsParametro!lRapido), False, RsParametro!lRapido)
   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   sClub = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   nPunto = IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto)
   lCierre = IIf(IsNull(RsParametro!lCierre), False, RsParametro!lCierre)
   nDecimal = IIf(IsNull(RsParametro!nDecimal), 2, RsParametro!nDecimal)
   nDias = IIf(IsNull(RsParametro!nDias), 2, RsParametro!nDias)
   nDiasDelivery = IIf(IsNull(RsParametro!nDiasDelivery), 2, RsParametro!nDiasDelivery)
   lEquivalencia = IIf(IsNull(RsParametro!lEquivalencia), False, RsParametro!lEquivalencia)
   lComboGeneral = IIf(IsNull(RsParametro!lComboGeneral), False, RsParametro!lComboGeneral)
   nSalir = 0
   nTipoCambio = 0
   multiLocal = IIf(IsNull(RsParametro!lmultilocal), False, RsParametro!lmultilocal)
   lImpresionDiaContable = IIf(IsNull(RsParametro!lImprimeDiaContable), False, RsParametro!lImprimeDiaContable)
   tContribuyenteEspecial = Trim(IIf(IsNull(RsParametro!tContribuyenteEspecial), "", RsParametro!tContribuyenteEspecial))
   fContribuyenteEspecial = Trim(IIf(IsNull(RsParametro!fContribuyenteEspecial), 2, RsParametro!fContribuyenteEspecial))
   
   
   
   'agenteretencion
   tTextoAgenteRetencion = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   'LOG
   lLogCajaRapida = IIf(IsNull(RsParametro!lLogCajaRapida), False, RsParametro!lLogCajaRapida)
   
   'consulta descargo al cierre
   lActivaConsultaDescargo = IIf(IsNull(RsParametro!lActivaConsultaDescargo), False, RsParametro!lActivaConsultaDescargo)
    
   'FORMATO VARIABLE
   nItemV = IIf(IsNull(RsParametro!nItemV), 0, RsParametro!nItemV)
   nCabeceraV = IIf(IsNull(RsParametro!nCabeceraV), 0, RsParametro!nCabeceraV)
   nPieV = IIf(IsNull(RsParametro!nPieV), 0, RsParametro!nPieV)
    
   'KDS
   lKDS = IIf(IsNull(RsParametro!lKDS), False, RsParametro!lKDS)
   sOrderInfo = IIf(IsNull(RsParametro!tOrderInfo), "", RsParametro!tOrderInfo)

   'club
   lClub = IIf(IsNull(RsParametro!lClub), False, RsParametro!lClub)

   'DIA contable
   lDiaContable = IIf(IsNull(RsParametro!lDiaContableAutomatico), True, RsParametro!lDiaContableAutomatico) ' true =automatico ' false=manual
   tHoraCierreDiaContable = IIf(IsNull(RsParametro!tHoraCierreDiaContable), "06:00", RsParametro!tHoraCierreDiaContable)


   'FACTURACION ELECTRONICA
   lFacturacionE = IIf(IsNull(RsParametro!lFacturacionE), False, RsParametro!lFacturacionE)
   tCodigoFE = IIf(IsNull(RsParametro!tCodigoFE), "000", RsParametro!tCodigoFE)
   tPieDocumento1 = IIf(IsNull(RsParametro!tPieDocumento1), " ", RsParametro!tPieDocumento1)
   lAmbienteProduccion = IIf(IsNull(RsParametro!lAmbienteFE), False, RsParametro!lAmbienteFE)
   RutaImgFE = IIf(IsNull(RsParametro!tRutaFE), "c:\", RsParametro!tRutaFE)

   'HUELLA
   lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
   lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)

   lImpresionCodigoBarras = IIf(IsNull(RsParametro!lImprimeCodigoBarras), False, RsParametro!lImprimeCodigoBarras)
   
   lEnvioAutomatico = IIf(IsNull(RsParametro!lEnvioAutomatico), False, RsParametro!lEnvioAutomatico)
   
   lEnvioProduccionUsuario = IIf(IsNull(RsParametro!lControlEnviosProduccion), False, RsParametro!lControlEnviosProduccion)
   
   lFEOfisis = IIf(IsNull(RsParametro!lFEOfisis), False, RsParametro!lFEOfisis)

   lPagoAntesImpresion = IIf(IsNull(RsParametro!lPagoAntesImpresion), False, RsParametro!lPagoAntesImpresion)
   
   If lPagoAntesImpresion Then
    lGlosaTrans = IIf(IsNull(RsParametro!lGlosaTransGratuita), "", RsParametro!lGlosaTransGratuita)
   End If
   ' COVER ECUADOR
    lcover = IIf(IsNull(RsParametro!lActivaCover), False, RsParametro!lActivaCover)
    sMontoMinCover = IIf(IsNull(RsParametro!tMontoMinCover), 0, RsParametro!tMontoMinCover)
    sCodItemCover = IIf(IsNull(RsParametro!tCodItemCover), "", RsParametro!tCodItemCover)

   '---------------------------------
    
    lNcOfisis = IIf(IsNull(RsParametro!lNcOfisis), 0, RsParametro!lNcOfisis)
   
   '----------------------------------
   ' notas de credito
    lactivaFechaNC = IIf(IsNull(RsParametro!lactivaFechaNC), False, RsParametro!lactivaFechaNC)
    lParcialNC = IIf(IsNull(RsParametro!lParcialNC), False, RsParametro!lParcialNC)
    lNCElimina = IIf(IsNull(RsParametro!lNCElimina), False, RsParametro!lNCElimina)
    lNCAnula = IIf(IsNull(RsParametro!lanula), False, RsParametro!lanula)
    'paperlees
    IPpape = Trim(LeerIni(App.Path + "\INFOREST.INI", "TCPIP", "IP", ""))
    PUERTOpape = Trim(LeerIni(App.Path + "\INFOREST.INI", "TCPIP", "PUERTO", ""))
    PUERTOLOCALpape = Trim(LeerIni(App.Path + "\INFOREST.INI", "TCPIP", "PUERTOLOCAL", ""))
    
   'TCANALVENTA
   sBoton1 = Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
   sBoton2 = Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
   sBoton3 = Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
   sBoton4 = Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
   sBoton5 = Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)
   obtieneVencimientoConexiones
   '----- visor 8" -----
   lvisor = IIf(IsNull(RsParametro!lvisor8), False, RsParametro!lvisor8)
   '------------------
   lQRFE = IIf(IsNull(RsParametro!lCodigoQrFE), 0, IIf(RsParametro!lCodigoQrFE = True, 1, 0))
   RutaImgFE = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE)
   
   lSAP = IIf(IsNull(RsParametro!lInteSAP), 0, IIf(RsParametro!lInteSAP = True, 1, 0))
   
   'FE Paperlees
   lFEpape = IIf(IsNull(RsParametro!lFEpape), False, RsParametro!lFEpape)

   'FE Spring
   lFESpring = IIf(IsNull(RsParametro!lFESpring), False, RsParametro!lFESpring)
   
   'FE Carvajal
   lFECarbajal = IIf(IsNull(RsParametro!lFECarbajal), False, RsParametro!lFECarbajal)
   
   lFEBiz = IIf(IsNull(RsParametro!lFEBiz), False, RsParametro!lFEBiz)
   
   If Not lFECarbajal Then
       'FE Carvajal
        lFECarbajal = IIf(IsNull(RsParametro!lFETCI), False, RsParametro!lFETCI)
   End If
   
   lDesPagoCheque = IIf(IsNull(RsParametro!lDesPagoCheque), False, RsParametro!lDesPagoCheque)
   lDesPagoOtra = IIf(IsNull(RsParametro!lDesPagoOtro), False, RsParametro!lDesPagoOtro)
   lDesactivaNCFP = IIf(IsNull(RsParametro!lDesactivaNCFP), False, RsParametro!lDesactivaNCFP)
   lActivaAnticipo = IIf(IsNull(RsParametro!lActivaAnticipo), False, RsParametro!lActivaAnticipo)
   lFeGoodHope = IIf(IsNull(RsParametro!lFEGood), False, RsParametro!lFEGood)
   
   lFEubl21 = IIf(IsNull(RsParametro!lFEubl21), False, RsParametro!lFEubl21)
   lBloqInafecto = IIf(IsNull(RsParametro!lBloqInafecto), False, RsParametro!lBloqInafecto)
   
   'lFEEcuador = IIf(IsNull(RsParametro!lEstupendoFE), False, RsParametro!lEstupendoFE)
   
   If IIf(IsNull(RsParametro!lEstupendoFE), False, RsParametro!lEstupendoFE) = True Then
    lFEEcuador = True
   End If
   
   If Not lFEBiz Then
    lFEBiz = lFeGoodHope
   End If
   
   lFEGesa = IIf(IsNull(RsParametro!lFEGesa), False, RsParametro!lFEGesa)

   sLey1 = IIf(IsNull(RsParametro!tLey1), "", RsParametro!tLey1)
   sValorLey1 = IIf(IsNull(RsParametro!tValorLey1), "", RsParametro!tValorLey1)
   If sValorLey1 = "" Or Not IsNumeric(sValorLey1) Then
    sValorLey1 = "0.00"
   End If
   
   lRutaImagenIntegraciones = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE) ' "\\192.168.3.37\compartir\IMAGENESQR"
   ' ELDC: 11/01/2021
   ' Se agrego y enlazo la variable nPorcentajePropina
   nPorcentajePropina = IIf(IsNull(RsParametro!nPorcentajePropina), 0, RsParametro!nPorcentajePropina)
   
   
   
   
   'FACTURACION_E_PERU
   On Error GoTo FacturacionIni
   
   If lFacturacionE And lFEpape = False And lFESpring = False Then
            Set CnFE = New Connection
            CnFE.Provider = "SQLOLEDB"
            CnFE.CursorLocation = adUseServer
            CnFE.ConnectionString = "User ID=" & sUserName & _
            ";password=" & sUserPassword & _
            ";Data Source=" & sRutaFE & _
            ";Initial Catalog=" & sMDBFE
            CnFE.CommandTimeout = 250
            CnFE.Open
    End If
   
FacturacionIni:
    If err.Number = "-2147467259" Then
       MsgBox "Facturacion Electronica, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    End If
    '-------------------------------
   
   
   
   
   On Error GoTo AlmacenIni
   Set CnAlmacen = New Connection
   Set CnAlmacenRemoto = New Connection
   Set cnAlmacenDefault = New Connection
   If lAlmacen Then
      ' Configuracion ini
      sAlmacenRuta = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
      sAlmacenMDB = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
      sLocal = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Configuracion", "LOCAL", "001"))
      
     
      CnAlmacen.Provider = "SQLOLEDB"
      CnAlmacen.CursorLocation = adUseServer
      CnAlmacen.ConnectionString = "User ID=" & sUserName & _
                                  ";password=" & sUserPassword & _
                                  ";Data Source=" & sAlmacenRuta & _
                                  ";Initial Catalog=" & sAlmacenMDB
       CnAlmacen.Open
       
       
       nLongitudAlmacen = Calcular("select ISNULL(AVG(LEN(TCODIGOLOCAL)),0) as codigo from TLOCALIDADES", CnAlmacen)
       
        If Len(sLocal) <> nLongitudAlmacen Then
        MsgBox "Error en la configuración del Local en el archivo ALMACEN.INI"
        Exit Sub
        End If
      
      
      
      
       Set cnAlmacenDefault = CnAlmacen
       localConectado = Calcular("select isnull(tresumido,'0') as codigo from vlocalidades where ip='" & sRuta & "' and bdinf='" & sMDB & "'", CnAlmacen)
       Dim verificaAlmacenRemoto As String
       verificaAdministracionCentralizada = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "REMOTO", "OFF"))
       If verificaAdministracionCentralizada = "ON" Then
            lAlmacenRemoto = True
            sRutaAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "SERVIDOR", "LOCAL"))
            sMDBAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "BASEDATO", "ALMACEN"))
       End If
       
   End If

    If localConectado = "0" Or localConectado = "" Then
            localConectado = sRazonComercial
    End If
   
   On Error GoTo InfhotelIni
   If lInfhotel Then
      sInfhotelRuta = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "SERVIDOR", "."))
      sInfhotelMDB = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "BASEDATO", "INFHOTEL"))
      sCajaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "CAJA", "01"))
      sPuntoVentaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "PUNTOVENTA", "01"))
      
      Set CnInfhotel = New Connection
      CnInfhotel.Provider = "SQLOLEDB"
      CnInfhotel.CursorLocation = adUseServer
      CnInfhotel.ConnectionString = "User ID=" & sUserName & _
                                    ";password=" & sUserPassword & _
                                    ";Data Source=" & sInfhotelRuta & _
                                    ";Initial Catalog=" & sInfhotelMDB
      CnInfhotel.Open
      sHotel = Calcular("select tHotel as Codigo from vCaja where tCaja='" & sCajaInfhotel & "'", CnInfhotel)
      sHotel = IIf(sHotel = "0", "01", sHotel)
      'sDireccion = Calcular("select tDireccion as Codigo From THOTEL where tHotel='" & sHotel & "'", CnInfhotel)
   End If
    
  'Call ValidaHARDkey 'OO
    Select Case pais
        Case "001"
            verificaFechaCodigoControl  ' avisa si faltan menos de 11 dias para el vencimiento de fecha de control
            textoComprobanteBolivia = "La reproducción total o parcial y/o  el uso no autorizado de esta Nota Fiscal, constituye un delito a ser sancionado conforme a Ley"
    End Select
        
   'frmFlash.Label5.Caption = "Módulo de Punto de Venta"
   'frmFlash.Show vbModal
          
   'Proceso de Caja
   Set RsCaja = Lib.OpenRecordset("select * from TCAJA", Cn)
   If RsCaja.RecordCount <> 0 Then
      RsCaja.MoveFirst
      RsCaja.Find ("tCaja='" & sCaja & "'")
      If RsCaja.EOF Then
         MsgBox "Error Faltal: No existe Caja Configurada", vbCritical, sMensaje
         End
      Else
         sPreCuenta = IIf(IsNull(RsCaja!tPrecuenta), "001", RsCaja!tPrecuenta)
         wComanda = IIf(IsNull(RsCaja!lComanda), False, RsCaja!lComanda)
         vComanda = IIf(IsNull(RsCaja!vComanda), False, RsCaja!vComanda)
         lEliminaC = IIf(IsNull(RsCaja!lMotivoEliminaC), False, RsCaja!lMotivoEliminaC)
         lElimina = IIf(IsNull(RsCaja!lMotivoElimina), False, RsCaja!lMotivoElimina)
         lComboPrecuenta = IIf(IsNull(RsCaja!lComboPrecuenta), False, RsCaja!lComboPrecuenta)
         lComboDocumento = IIf(IsNull(RsCaja!lComboDocumento), False, RsCaja!lComboDocumento)
         lPasswordC = IIf(IsNull(RsCaja!lPasswordC), False, RsCaja!lPasswordC)
         lPassword = IIf(IsNull(RsCaja!lPassword), False, RsCaja!lPassword)
         sGrupoDefault = IIf(IsNull(RsCaja!tgrupo), "01", RsCaja!tgrupo)
         lConsumo1 = IIf(IsNull(RsCaja!lConsumo1), False, RsCaja!lConsumo1)
         lConsumo2 = IIf(IsNull(RsCaja!lConsumo2), False, RsCaja!lConsumo2)
         lConsumo3 = IIf(IsNull(RsCaja!lConsumo3), False, RsCaja!lConsumo3)
         lPrecuentaImpresora = IIf(IsNull(RsCaja!lPrecuenta), False, RsCaja!lPrecuenta)
         lAdicion = IIf(IsNull(RsCaja!lAdicion), False, RsCaja!lAdicion)
         lPrecuentaAgrupada = IIf(IsNull(RsCaja!lPrecuentaAgrupada), False, RsCaja!lPrecuentaAgrupada)
         sTipoPedidoPD = IIf(IsNull(RsCaja!tTipoPedido), "01", RsCaja!tTipoPedido)
         lObligaPrinter = IIf(IsNull(RsCaja!lObligaPrinter), False, RsCaja!lObligaPrinter)
         lObligaPrecuenta = IIf(IsNull(RsCaja!lObligaPrecuenta), False, RsCaja!lObligaPrecuenta)
         lObligaCierre = IIf(IsNull(RsCaja!lObligaCierre), False, RsCaja!lObligaCierre)
         lFiltroTipoPedido = IIf(IsNull(RsCaja!lFiltroTipoPedido), False, RsCaja!lFiltroTipoPedido)
         nPuerto = IIf(IsNull(RsCaja!nPuerto), 0, RsCaja!nPuerto)
         nBalanzaPuerto = IIf(IsNull(RsCaja!nBalanzaPuerto), 0, RsCaja!nBalanzaPuerto)
         tMensaje1 = Trim(IIf(IsNull(RsCaja!tMensaje1), "", RsCaja!tMensaje1))
         tMensaje2 = Trim(IIf(IsNull(RsCaja!tMensaje2), "", RsCaja!tMensaje2))
         lCancelacion = IIf(IsNull(RsCaja!lCancelacion), False, RsCaja!lCancelacion)
         lCambioMesa = IIf(IsNull(RsCaja!lCambioMesa), False, RsCaja!lCambioMesa)
         lDirecto = IIf(IsNull(RsCaja!lDirecto), False, RsCaja!lDirecto)
         lVisaNet = IIf(IsNull(RsCaja!lVisaNet), False, RsCaja!lVisaNet)
         lImpuestoPrecuenta = IIf(IsNull(RsCaja!lImpuestoPrecuenta), False, RsCaja!lImpuestoPrecuenta)
         lDocumentoAgrupado = IIf(IsNull(RsCaja!lDocumentoAgrupado), False, RsCaja!lDocumentoAgrupado)
         lOrden = IIf(IsNull(RsCaja!lOrden), False, RsCaja!lOrden)
         lValorCortesia = IIf(IsNull(RsCaja!lValorCortesia), False, RsCaja!lValorCortesia)
         lObservacion = IIf(IsNull(RsCaja!lObservacion), False, RsCaja!lObservacion)
         lCajaRapida = IIf(IsNull(RsCaja!lCajaRapida), False, RsCaja!lCajaRapida)
         lPropiedadDocumento = IIf(IsNull(RsCaja!lPropiedadDocumento), False, RsCaja!lPropiedadDocumento)
         lPropiedadPrecuenta = IIf(IsNull(RsCaja!lPropiedadPrecuenta), False, RsCaja!lPropiedadPrecuenta)
         lPrecioNetoPrecuenta = IIf(IsNull(RsCaja!lPrecioNetoPrecuenta), False, RsCaja!lPrecioNetoPrecuenta)
         nLimitePrecuenta = IIf(IsNull(RsCaja!nLimitePrecuenta), 0, RsCaja!nLimitePrecuenta)
         nLimiteReimpresion = IIf(IsNull(RsCaja!nLimiteReimpresion), 0, RsCaja!nLimiteReimpresion)
         lPasswordTransferencia = IIf(IsNull(RsCaja!lPasswordTransferencia), False, RsCaja!lPasswordTransferencia)
         sUnidadNegocio = IIf(IsNull(RsCaja!tUnidadNegocio), "01", RsCaja!tUnidadNegocio)
         
         CD = IIf(IsNull(RsCaja!lCD), False, RsCaja!lCD)
         WEBAP = IIf(IsNull(RsCaja!lWebAp), False, RsCaja!lWebAp)
         MESA247 = IIf(IsNull(RsCaja!lMesa247), False, RsCaja!lMesa247)
         EAN13 = IIf(IsNull(RsCaja!EAN13), False, RsCaja!EAN13)
         lBuscaPedidoNumero = IIf(IsNull(RsCaja!lBuscaPedidoNumero), False, RsCaja!lBuscaPedidoNumero)
         lMultiCajero = IIf(IsNull(RsCaja!lMultiCajero), False, RsCaja!lMultiCajero)
         lMCPV = IIf(IsNull(RsCaja!lMCPV), False, RsCaja!lMCPV)
         lFechaEntregaDelivery = IIf(IsNull(RsCaja!lFechaEntregaDelivery), False, RsCaja!lFechaEntregaDelivery)
         lCCVOX = IIf(IsNull(RsCaja!lCCVOX), False, RsCaja!lCCVOX)
         lObservacionDocumento = IIf(IsNull(RsCaja!lObservacionDocumento), False, RsCaja!lObservacionDocumento)
         lObservacionCabDoc = IIf(IsNull(RsCaja!lObservacionCabDoc), False, RsCaja!lObservacionCabDoc)
         lImprimeMotivoDescuentoFB = IIf(IsNull(RsCaja!lMotivoDescuento), False, RsCaja!lMotivoDescuento)
         lObservacionPrecuenta = IIf(IsNull(RsCaja!lObservacionPrecuenta), False, RsCaja!lObservacionPrecuenta)
         nLongitudBarra = Trim(IIf(IsNull(RsCaja!nLongitudBarra), 0, RsCaja!nLongitudBarra))
         lCapturaPeso = IIf(IsNull(RsCaja!lCapturaPeso), False, RsCaja!lCapturaPeso)
         lPagoRapidoPV = IIf(IsNull(RsCaja!lPagoRapidoPV), False, RsCaja!lPagoRapidoPV)
         tTextoConsumo = IIf(IsNull(RsCaja!tTextoConsumo), False, RsCaja!tTextoConsumo)
         lDisgrega = IIf(IsNull(RsCaja!lDisgrega), False, RsCaja!lDisgrega)
         lPasswordImportarPedido = IIf(IsNull(RsCaja!lPasswordImportarPedido), False, RsCaja!lPasswordImportarPedido)
         lImpreEquivaPrecuenta = IIf(IsNull(RsCaja!lequivadolaprecuenta), False, RsCaja!lequivadolaprecuenta)
         lDescripcionAlternativa = IIf(IsNull(RsCaja!lActivaImpDscAlternativa), False, RsCaja!lActivaImpDscAlternativa)
         lPagoRapido = IIf(IsNull(RsCaja!lPagoRapido), False, RsCaja!lPagoRapido)
         lPasswordPorCobrar = IIf(IsNull(RsCaja!lPasswordPorCobrar), False, RsCaja!lPasswordPorCobrar)
         lmodificatipoPedido = IIf(IsNull(RsCaja!lmodificatipoPedido), False, RsCaja!lmodificatipoPedido)
         lSiab = IIf(IsNull(RsCaja!lSiab), False, RsCaja!lSiab)
         'codigobarra
         lCodigoReciboIngreso = IIf(IsNull(RsCaja!lCodigoReciboIngreso), False, RsCaja!lCodigoReciboIngreso)
        
         lImprimeImagCabPrecuenta = IIf(IsNull(RsCaja!lImprimeImagCabPrecuenta), False, RsCaja!lImprimeImagCabPrecuenta)
         lImprimeImagPiePrecuenta = IIf(IsNull(RsCaja!lImprimeImagPiePrecuenta), False, RsCaja!lImprimeImagPiePrecuenta)
         lAccesoDespachoPedido = IIf(IsNull(RsCaja!lAccesoDespachoPedido), False, RsCaja!lAccesoDespachoPedido)
        
         lBuscarPedidoVisualizarGrilla = IIf(IsNull(RsCaja!lBuscarPedidoVisualizarGrilla), False, RsCaja!lBuscarPedidoVisualizarGrilla)
         lBuscarPedidoFiltrarMesa = IIf(IsNull(RsCaja!lBuscarPedidoFiltrarMesa), False, RsCaja!lBuscarPedidoFiltrarMesa)
        
         lBloqueaPrecuenta = IIf(IsNull(RsCaja!lBloqueaPrecuenta), False, RsCaja!lBloqueaPrecuenta)

         lRotulado = IIf(IsNull(RsCaja!lRotulado), False, RsCaja!lRotulado)
                
         lMultiAreaSubGrupo = IIf(IsNull(RsCaja!lMultiAreaSubGrupo), False, RsCaja!lMultiAreaSubGrupo)
         
         lMultiAreaCaja = IIf(IsNull(RsCaja!lMultiAreaCaja), False, RsCaja!lMultiAreaCaja)

         lConsumo4 = IIf(IsNull(RsCaja!lConsumo4), False, RsCaja!lConsumo4)
         
         lPrecuentaNoValorizada = IIf(IsNull(RsCaja!lPrecuentaNoValorizada), False, RsCaja!lPrecuentaNoValorizada)
         
         lEnvioProduccionCaja = IIf(IsNull(RsCaja!lClaveEnvioProduccion), False, RsCaja!lClaveEnvioProduccion)
         lImpPropina = IIf(IsNull(RsCaja!lImpPropina), False, RsCaja!lImpPropina)
         lImpComandaf2 = IIf(IsNull(RsCaja!lImpComandaf2), False, RsCaja!lImpComandaf2)
         lPassOtrosPagos = IIf(IsNull(RsCaja!lPassOtrosPagos), False, RsCaja!lPassOtrosPagos)
         lImpClienteCabDoc = IIf(IsNull(RsCaja!lImpClienteCab), False, RsCaja!lImpClienteCab)
         lBuscarPedidoBarraMesero = IIf(IsNull(RsCaja!lBuscarPedidoBarra), False, RsCaja!lBuscarPedidoBarra)
         lIntVisaNetQR = IIf(IsNull(RsCaja!lIntVisaNetQR), False, RsCaja!lIntVisaNetQR)
         lIntMercadoPagoQR = IIf(IsNull(RsCaja!lIntMercadoPagoQR), False, RsCaja!lIntMercadoPagoQR)
         lIntMercadoPagoQREstatico = IIf(IsNull(RsCaja!lIntMercadoPagoQREstatico), False, RsCaja!lIntMercadoPagoQREstatico)
         lImpComandaBarra = IIf(IsNull(RsCaja!lImpComandaBarra), False, RsCaja!lImpComandaBarra)
         lbloqueadocconsumo = IIf(IsNull(RsCaja!lbloqueadocconsumo), False, RsCaja!lbloqueadocconsumo)
         lCashDro = IIf(IsNull(RsCaja!lCashDro), False, RsCaja!lCashDro)
         lBloqueoAfectoInafecto = IIf(IsNull(RsCaja!lBloqueoAfectoInafecto), False, RsCaja!lBloqueoAfectoInafecto)
         lIntPatio = IIf(IsNull(RsCaja!lIntPatio), False, RsCaja!lIntPatio)
         lIntPagoEfectivo = IIf(IsNull(RsCaja!lIntPagoEfectivo), False, RsCaja!lIntPagoEfectivo)
         lPagoDirectoCD = IIf(IsNull(RsCaja!lPagoDirectoCD), False, RsCaja!lPagoDirectoCD)
         
         ' ELDC: 11/01/2021
         ' Se agrego y enlazo la variable lImprimePropinaSugerida
         lImprimePropinaSugerida = IIf(IsNull(RsCaja!lImprimePropinaSugerida), False, RsCaja!lImprimePropinaSugerida)
         lActivarAvisoRetenciones = IIf(IsNull(RsCaja!lActivarAvisoRetenciones), False, RsCaja!lActivarAvisoRetenciones)
         
         
         TextoPagoEfectivo = ""
         If lIntPagoEfectivo = True Then
            TextoPagoEfectivo = "- Transferencias bancarias vía PagoEfectivo: Paga en BBVA,BCP, Interbank, Scotiabank, BanBif, Caja Arequipa,Pichincha, a través de la banca por internet o banca móvil en la opción pago de servicios."
         End If
         
         On Error GoTo ex
            Dim RsBalanza As Recordset
            Isql = "Select * from vbalanza where tcaja='" & sCaja & "'"
            Set RsBalanza = Lib.OpenRecordset(Isql, Cn)
            If Not (RsBalanza.EOF Or RsBalanza.BOF) Then
                RsBalanza.MoveFirst
                With RsBalanza
                    nBalanzaBS = CLng(IIf(IsNull(!nDato1), 0, !nDato1))
                    nBalanzaBD = CInt(IIf(IsNull(!nDato2), 0, !nDato2))
                    nBalanzaParidad = LCase(Left(IIf(IsNull(!nDato3) Or Trim(!nDato3) = "", "n", !nDato3), 1))
                    nBalanzaBP = CDbl(IIf(IsNull(!nDato4), 0, !nDato4))
                    nBalanzaCF = IIf(IsNull(!nDato5), "", !nDato5)
                    nBalanzaPuerto = CInt(IIf(IsNull(!nDato6), 0, !nDato6))
                    nBalanzaComando = IIf(IsNull(!nDato7), "", !nDato7)
                    lBalTiempo = CInt(IIf(IsNull(!nDato8), 0, !nDato8))
                    lBal = IIf(IsNull(!lActivo), False, !lActivo)
                End With
            Else
                 nBalanzaComando = ""
                 nBalanzaPuerto = 0
                 nBalanzaBS = 0
                 nBalanzaBD = 0
                 nBalanzaParidad = ""
                 nBalanzaBP = 0
                 lBalTiempo = 0
                 nBalanzaCF = ""
            End If
ex:
         'TVS=======================
         'lCompatibilidadTVS = IIf(IsNull(RsCaja!lCompatibilidadTVS), False, RsCaja!lCompatibilidadTVS)
         'If (lCompatibilidadTVS) Then
         '   Call CrearIniTVS
         'End If
      End If
   Else
      MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   'Proceso de Correlativo
   Isql = "select * from vTipoDocumento where Descripcion='FACTURA'"
   Set RsBusca = Lib.OpenRecordset(Isql, Cn)
   If RsBusca.RecordCount > 0 Then
      Isql = "select * from TTIPODOCUMENTOIMPRESORA where tCaja ='" & sCaja & "' and tTipoEmision='" & RsBusca!codigo & "'"
      Set RsFactura = Lib.OpenRecordset(Isql, Cn)
      If RsFactura.RecordCount > 0 Then
         NFactura = RsFactura!tUltimoNumero
      Else
         NFactura = "Sin Correlativo"
      End If
   Else
     NFactura = "Sin Correlativo"
   End If
   
   Set RsCaja = Nothing
   Set RsParametro = Nothing
   Set RsBusca = Nothing
   Set RsFactura = Nothing
   
   
   Isql = "select *  from TTIPODOCUMENTOIMPRESORA where ttipoemision<>'00' and lFacturacionElectronica=1 and tCaja ='" & sCaja & "'"
   Set RsFpape = Lib.OpenRecordset(Isql, Cn)
    If lFEpape Then
        If RsFpape.RecordCount > 0 Then
            If FacturarTCPIP(1, "", 0) = False Then
                End
            End If
        End If
    End If
   
  
   
   wInicio = False
   If lMCPV Then
    mdiPuntoVenta.Show
   Else
    frmAcceso.Caption = "Inforest Módulo de Punto de Venta v." & App.Major & "." & App.Minor & "." & App.Revision
    frmAcceso.Show vbModal
    If wEnter = True Then
      mdiCajaRapida.Show
    End If
   End If
   Exit Sub
   
InforestIni:
    If err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
AlmacenIni:
    If err.Number = "-2147467259" Then
       MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
InfhotelIni:
    If err.Number = "-2147467259" Then
       MsgBox "Infhotel, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
End Sub

Public Sub ActivaInicio(Activa As Boolean)
    mdiCajaRapida.cmdOpcion(1).Enabled = Activa
    mdiCajaRapida.cmdOpcion(2).Enabled = Activa
    mdiCajaRapida.mnuVenta.Enabled = Activa
    mdiCajaRapida.mnuCierre.Enabled = Activa
End Sub
