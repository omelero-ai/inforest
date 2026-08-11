Attribute VB_Name = "modAdicion"
 
Public nRefresca As Long
Public nSalir As Long
Public hk As New License
 
 Sub Main()
  Dim existe As Boolean
  Dim AnoMesSv As String
  Dim directorioSistema  As String
  Dim archI As String
 
    lHARDkey = False
    '---------------
    If App.PrevInstance Then
        MsgBox "Ya se esta ejecutando el Aplicativo!", vbInformation, "Atención!!!"
        End
    End If
    '------------
    lVersionEducativa = False
    
    Open App.Path & "\TIEMPO.INI" For Input As #1   ' Abre el archivo para recibir los datos.
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
       Input #1, nRefresca, nSalir
    Loop
    Close #1   ' Cierra el archivo.
    
    On Error GoTo InforestIni
    
    '==========================================================
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
    
       Dim Tamano  As String
    Tamano = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SIZEPIEPRECUENTA", "0"))
    
    If Tamano = "" Or Tamano = "0" Or Not IsNumeric(Tamano) Then
        SizeTamanoPiePrecuenta = 0
    Else
        SizeTamanoPiePrecuenta = CInt(Tamano)
    End If
    
    
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    SYPROV = Trim(LeerIni(App.Path + "\INFOREST.INI", "CONEXION", "SYPROV", "SAOLEDB"))
    AvisoLicencia = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "AVISO", 1))
    
    Dim RsParametro As Recordset
    Dim RsTc As Recordset
    Dim RsCaja As Recordset
        
    Set Lib = New Libreria16.Applications
    
    sUserName = "infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    
    'auditoria
    nCorrelativoAcceso = 0
    tModuloSeg = "02" 'XX= CODIGO DEL MODULO DE LA TABLA DE MMODULO DE SEGURIDAD
    Set CnSeg = New Connection
      
    CnSeg.Provider = "SQLOLEDB"
    CnSeg.CursorLocation = adUseServer
    CnSeg.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=INFSEGURIDAD"
    CnSeg.CommandTimeout = 250
    
    If xlCentral = True Then
        '========================================== inforest central GCAA 27/02/2020
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



  '====================================================
  If lVersionEducativa = True Then
     directorioSistema = ObtenerDirectorioSO
     directorioSistema = directorioSistema & "\Infhotel.csg"
     existe = FileExists(directorioSistema)
     If existe = True Then
        AnoMesSv = Calcular("select dbo.returnAnoMes(getdate()) as codigo", Cn)
        archI = obtieneAnoMes(directorioSistema)
        If Val(AnoMesSv) > Val(archI) Then
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
    
    
    'Diccionario
    
    Mesero = DiccionarioValor("Mesero")
    
  
    '===================================================
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "ADICION"
   
   sRUC = SoloNumeros(IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria))
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

    If Not GetFileTimes(mstrFile, dteCreate, dteAccess, dteModify, True) Then
        Call MsgBox("Error getting file dates.", vbCritical, "Unable to Read File Dates")
        End
    End If

    If Not GetFileTimes(sActRuta, dteCreateACT, dteAccessACT, dteModifyACT, True) Then
        'Call MsgBox("Error getting file dates.", vbCritical, "Unable to Read File Dates")
        'End
        dteModifyACT = dteModify
    End If

     
    sVersion = ""
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    sVersion = Calcular("SELECT tVersion As Codigo FROM tParametro", Cn)
    
    If sVersion <> sVersionExe Or (sVersion = sVersionExe And dteModifyACT > dteModify) Then
        MsgBox "Existe Una Nueva Version Disponible", vbInformation, sMensaje
    
        'CREA LA CARPETA PARA BACKUP DE EXES
        Dim Backup As String
        Backup = App.Path + "\ExesHistoricos"
    
        'CREA LA CARPETA Y VALIDA QUE NO EXISTA
        If Dir(Backup, vbDirectory) = "" Then
           MkDir (Backup)
        End If
        
        Call IniciarActualizador("ADICION")
        'Shell App.Path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
        End
    Else
         Call IniciarActualizador("INICIO")
        'Sleep (3000)
    End If
       
    '------
    
    'validacion acceso Licencias
     validacionLicenciasInfhotel
    
    '===================================================================
    '===================================================================

   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
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
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sPie = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
    sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   sMail = IIf(IsNull(RsParametro!tEmail), "", RsParametro!tEmail)
   sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tpassword), "", RsParametro!tpassword)))
   sElimina = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
   nFItem = IIf(IsNull(RsParametro!nItem), 0, RsParametro!nItem)
   lPrinter = IIf(IsNull(RsParametro!lPrinter), False, RsParametro!lPrinter)
 
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   lRapido = IIf(IsNull(RsParametro!lRapido), False, RsParametro!lRapido)
   'HUELLA
   lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
   lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)
   'KDS
   nPuerto = 0
   sPiePreCuenta = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   lComboGeneral = IIf(IsNull(RsParametro!lComboGeneral), False, RsParametro!lComboGeneral)
   'KDS
   lKDS = IIf(IsNull(RsParametro!lKDS), False, RsParametro!lKDS)
   sOrderInfo = IIf(IsNull(RsParametro!tOrderInfo), "", RsParametro!tOrderInfo)

   'DIA contable
   lDiaContable = IIf(IsNull(RsParametro!lDiaContableAutomatico), True, RsParametro!lDiaContableAutomatico) ' true =automatico ' false=manual
   tHoraCierreDiaContable = IIf(IsNull(RsParametro!tHoraCierreDiaContable), "06:00", RsParametro!tHoraCierreDiaContable)
   
   'club
    lClub = IIf(IsNull(RsParametro!lClub), False, RsParametro!lClub)

   'consulta descargo al cierre
   lActivaConsultaDescargo = IIf(IsNull(RsParametro!lActivaConsultaDescargo), False, RsParametro!lActivaConsultaDescargo)

   'agenteretencion
   tTextoAgenteRetencion = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   
   lEnvioAutomatico = IIf(IsNull(RsParametro!lEnvioAutomatico), False, RsParametro!lEnvioAutomatico)
   
   'lcover = IIf(IsNull(RsParametro!lActivaCover), False, RsParametro!lActivaCover)
   
   lEnvioProduccionUsuario = IIf(IsNull(RsParametro!lControlEnviosProduccion), False, RsParametro!lControlEnviosProduccion)
   
   ' COVER ECUADOR
    lcover = IIf(IsNull(RsParametro!lActivaCover), False, RsParametro!lActivaCover)
    sMontoMinCover = IIf(IsNull(RsParametro!tMontoMinCover), 0, RsParametro!tMontoMinCover)
    sCodItemCover = IIf(IsNull(RsParametro!tCodItemCover), "", RsParametro!tCodItemCover)

    sLey1 = IIf(IsNull(RsParametro!tLey1), "", RsParametro!tLey1)
    sValorLey1 = IIf(IsNull(RsParametro!tValorLey1), "", RsParametro!tValorLey1)

   '-----------------
   lBloqInafecto = IIf(IsNull(RsParametro!lBloqInafecto), False, RsParametro!lBloqInafecto)
   
   
    lKDSInforest = IIf(IsNull(RsParametro!lKDSInforest), 0, RsParametro!lKDSInforest)
   lKDSxArea = IIf(IsNull(RsParametro!lKDSxArea), 0, RsParametro!lKDSxArea)
   lKDSTodos = IIf(IsNull(RsParametro!lKDSTodos), 0, RsParametro!lKDSTodos)
   
    lActivaAnticipo = IIf(IsNull(RsParametro!lActivaAnticipo), 0, RsParametro!lActivaAnticipo)
    
    

   'Canal de Venta
   sBoton1 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
   sBoton2 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
   sBoton3 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
   sBoton4 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
   sBoton5 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)

   obtieneVencimientoConexiones

   On Error GoTo AlmacenIni
   Set CnAlmacen = New Connection
   If lAlmacen Then
      ' Configuracion ini
      sAlmacenRuta = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
      sAlmacenMDB = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
      sLocal = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Configuracion", "LOCAL", "001"))
     
     On Error GoTo ValidaAlmacen
     
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
      
ValidaAlmacen:
        If err.Number = "-2147467259" Then
           'MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
            If MsgBox("Almacen, SQL Server no Encontrado, Falla de Conectividad!!! " & vbNewLine & " ¿Desea Continuar sin enlace a Almacen?" & vbNewLine & " de continuar con el proceso, se informa que los costos de los productos " & vbNewLine & " vendidos sin enlace a almacen no se grabaran!!! ", vbYesNo, "Almacen") = vbNo Then
                Exit Sub
            Else
                Call Log_Inforest("Almacen", "CONEXIONALMACEN", TipoIni, "Caja: " + sCaja, Environ("computername"), "Ingreso sin conexion de almacen a Inforest", "", "FALLA PROCESO", "INICIO")
                lAlmacen = False
            End If
        End If
        
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
         lPasswordC = IIf(IsNull(RsCaja!lPasswordC), False, RsCaja!lPasswordC)
         lPassword = IIf(IsNull(RsCaja!lPassword), False, RsCaja!lPassword)
         sGrupoDefault = IIf(IsNull(RsCaja!tgrupo), "01", RsCaja!tgrupo)
         lPrecuentaAgrupada = IIf(IsNull(RsCaja!lPrecuentaAgrupada), False, RsCaja!lPrecuentaAgrupada)
         lPrecuentaImpresora = IIf(IsNull(RsCaja!lPrecuenta), False, RsCaja!lPrecuenta)
         lObligaPrinter = IIf(IsNull(RsCaja!lObligaPrinter), False, RsCaja!lObligaPrinter)
         nPuerto = IIf(IsNull(RsCaja!nPuerto), 0, RsCaja!nPuerto)
         tMensaje1 = Trim(IIf(IsNull(RsCaja!tMensaje1), "", RsCaja!tMensaje1))
         tMensaje2 = Trim(IIf(IsNull(RsCaja!tMensaje2), "", RsCaja!tMensaje2))
         lDirecto = IIf(IsNull(RsCaja!lDirecto), False, RsCaja!lDirecto)
         lComboPrecuenta = IIf(IsNull(RsCaja!lComboPrecuenta), False, RsCaja!lComboPrecuenta)
         lCambioMesa = IIf(IsNull(RsCaja!lCambioMesa), False, RsCaja!lCambioMesa)
         lImpuestoPrecuenta = IIf(IsNull(RsCaja!lImpuestoPrecuenta), False, RsCaja!lImpuestoPrecuenta)
         lOrden = IIf(IsNull(RsCaja!lOrden), False, RsCaja!lOrden)
         lObservacion = IIf(IsNull(RsCaja!lObservacion), False, RsCaja!lObservacion)
         lCajaRapida = IIf(IsNull(RsCaja!lCajaRapida), False, RsCaja!lCajaRapida)
         lPropiedadPrecuenta = IIf(IsNull(RsCaja!lPropiedadPrecuenta), False, RsCaja!lPropiedadPrecuenta)
         lPropiedadDocumento = IIf(IsNull(RsCaja!lPropiedadDocumento), False, RsCaja!lPropiedadDocumento)
         nLimitePrecuenta = IIf(IsNull(RsCaja!nLimitePrecuenta), 0, RsCaja!nLimitePrecuenta)
         sUnidadNegocio = IIf(IsNull(RsCaja!tUnidadNegocio), "01", RsCaja!tUnidadNegocio)
         nLimiteReimpresion = IIf(IsNull(RsCaja!nLimiteReimpresion), 0, RsCaja!nLimiteReimpresion)
         lPasswordTransferencia = IIf(IsNull(RsCaja!lPasswordTransferencia), False, RsCaja!lPasswordTransferencia)
         lMCPV = IIf(IsNull(RsCaja!lMCPV), False, RsCaja!lMCPV)
         lObservacionDocumento = IIf(IsNull(RsCaja!lObservacionDocumento), False, RsCaja!lObservacionDocumento)
         lObservacionPrecuenta = IIf(IsNull(RsCaja!lObservacionPrecuenta), False, RsCaja!lObservacionPrecuenta)
         
         lImprimeImagCabPrecuenta = IIf(IsNull(RsCaja!lImprimeImagCabPrecuenta), False, RsCaja!lImprimeImagCabPrecuenta)
         lImprimeImagPiePrecuenta = IIf(IsNull(RsCaja!lImprimeImagPiePrecuenta), False, RsCaja!lImprimeImagPiePrecuenta)

         lAccesoDespachoPedido = IIf(IsNull(RsCaja!lAccesoDespachoPedido), False, RsCaja!lAccesoDespachoPedido)
         'marzo 2011
         lPasswordImportarPedido = IIf(IsNull(RsCaja!lPasswordImportarPedido), False, RsCaja!lPasswordImportarPedido)
         lDisgrega = IIf(IsNull(RsCaja!lDisgrega), False, RsCaja!lDisgrega)
         sTipoPedidoPD = IIf(IsNull(RsCaja!tTipoPedido), "01", RsCaja!tTipoPedido)
         lSiab = IIf(IsNull(RsCaja!lSiab), False, RsCaja!lSiab)
         
         lBusquedaReservas = IIf(IsNull(RsCaja!lBusquedaReservas), False, RsCaja!lBusquedaReservas)
                 
         lBuscarPedidoVisualizarGrilla = IIf(IsNull(RsCaja!lBuscarPedidoVisualizarGrilla), False, RsCaja!lBuscarPedidoVisualizarGrilla)
         lBuscarPedidoFiltrarMesa = IIf(IsNull(RsCaja!lBuscarPedidoFiltrarMesa), False, RsCaja!lBuscarPedidoFiltrarMesa)
         
         lBloqueaPrecuenta = IIf(IsNull(RsCaja!lBloqueaPrecuenta), False, RsCaja!lBloqueaPrecuenta)
         '---- MESA24/7
         MESA247 = IIf(IsNull(RsCaja!lMesa247), False, RsCaja!lMesa247)
         
                           ' nov 2010
         '==========================
         lImpreEquivaPrecuenta = IIf(IsNull(RsCaja!lequivadolaprecuenta), False, RsCaja!lequivadolaprecuenta)
         '==========================
         'PAGORAPIDO=======================
         lPagoRapido = IIf(IsNull(RsCaja!lPagoRapido), False, RsCaja!lPagoRapido)
          
          'passwordporcobrar / modificatipopedido
         lPasswordPorCobrar = IIf(IsNull(RsCaja!lPasswordPorCobrar), False, RsCaja!lPasswordPorCobrar)
         lmodificatipoPedido = IIf(IsNull(RsCaja!lmodificatipoPedido), False, RsCaja!lmodificatipoPedido)
         
                 
         lMultiAreaSubGrupo = IIf(IsNull(RsCaja!lMultiAreaSubGrupo), False, RsCaja!lMultiAreaSubGrupo)
         
         lMultiAreaCaja = IIf(IsNull(RsCaja!lMultiAreaCaja), False, RsCaja!lMultiAreaCaja)
         
         lPrecuentaNoValorizada = IIf(IsNull(RsCaja!lPrecuentaNoValorizada), False, RsCaja!lPrecuentaNoValorizada)
          
         lEnvioProduccionCaja = IIf(IsNull(RsCaja!lClaveEnvioProduccion), False, RsCaja!lClaveEnvioProduccion)
         
         lImpPropina = IIf(IsNull(RsCaja!lImpPropina), False, RsCaja!lImpPropina)
         lImpComandaf2 = IIf(IsNull(RsCaja!lImpComandaf2), False, RsCaja!lImpComandaf2)
         
         lBuscarPedidoBarraMesero = IIf(IsNull(RsCaja!lBuscarPedidoBarra), False, RsCaja!lBuscarPedidoBarra)
         'TVS=======================
         lCompatibilidadTVS = IIf(IsNull(RsCaja!lCompatibilidadTVS), False, RsCaja!lCompatibilidadTVS)
         
         lActivaTresDecimales = IIf(IsNull(RsCaja!lActivaTresDecimales), False, RsCaja!lActivaTresDecimales)
         
         lImprimePropinaSugerida = IIf(IsNull(RsCaja!lImprimePropinaSugerida), False, RsCaja!lImprimePropinaSugerida)
         nPorcentajePropina = IIf(IsNull(RsParametro!nPorcentajePropina), 0, RsParametro!nPorcentajePropina)

         lBloqueoAnulaPedido = IIf(IsNull(RsCaja!lBloqueoAnulaPedido), False, RsCaja!lBloqueoAnulaPedido)
         lBloqueoAnulaItemsPedido = IIf(IsNull(RsCaja!lBloqueoAnulaItemsPedido), False, RsCaja!lBloqueoAnulaItemsPedido)
         
        lBloqueoCanal1 = IIf(IsNull(RsCaja!lBloqueaCanal1), False, RsCaja!lBloqueaCanal1)
        lBloqueoCanal2 = IIf(IsNull(RsCaja!lBloqueaCanal2), False, RsCaja!lBloqueaCanal2)
        lBloqueoCanal3 = IIf(IsNull(RsCaja!lBloqueaCanal3), False, RsCaja!lBloqueaCanal3)
        lBloqueoCanal4 = IIf(IsNull(RsCaja!lBloqueaCanal4), False, RsCaja!lBloqueaCanal4)
        lBloqueoCanal5 = IIf(IsNull(RsCaja!lBloqueaCanal5), False, RsCaja!lBloqueaCanal5)
        
        lImprimePrecuentaCabunica = IIf(IsNull(RsCaja!lImprimePrecuentaCabunica), False, RsCaja!lImprimePrecuentaCabunica)
        lBloquearPedidoSinMesa = IIf(IsNull(RsCaja!lBloquearPedidoSinMesa), False, RsCaja!lBloquearPedidoSinMesa)
        lAutorizacionSinMeza = IIf(IsNull(RsCaja!lAutorizacionSinMeza), False, RsCaja!lAutorizacionSinMeza)
        lImprimePrepagoPrecuenta = IIf(IsNull(RsCaja!lImprimePagoPrecuenta), False, RsCaja!lImprimePagoPrecuenta)
        
         If (lCompatibilidadTVS) Then
            Call CrearIniTVS
         End If
         
         
      End If
   Else
      MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
      End
   End If
     
     
   'TIPO CAMBIO
   If pais = "002" Then
      RTipoCambio
'   Else
'            Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
'
'            If RsTc.EOF Then
'               nTC = 0
'            Else
'               nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
'            End If
'
'            Set RsTc = Nothing
'            Set RsCaja = Nothing
'            Set RsParametro = Nothing
'            wInicio = False
'            If nTC = 0 Then
'               MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
'               Exit Sub
'            End If
   End If
   
   
   Call ValidaHARDkey 'OO
   frmMozoUsuario.Show vbModal
Exit Sub
InforestIni:
    If err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    End If
    MsgBox "Archivo de configuracion con uno o mas parametros Incompletos o" & Chr(13) & "Incorrectos, comunicarse con Infhotel Servicio Informaticos S.A.C.", vbCritical + vbOKOnly
    Exit Sub
AlmacenIni:
    If err.Number = "-2147467259" Then
       MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    End If
    MsgBox "Archivo de configuracion con uno o mas parametros Incompletos o" & Chr(13) & "Incorrectos, comunicarse con Infhotel Servicio Informaticos S.A.C.", vbCritical + vbOKOnly
    MsgBox err.Description
    Exit Sub
InfhotelIni:
    If err.Number = "-2147467259" Then
       MsgBox "Infhotel, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
End Sub

Public Sub RTipoCambio()
   'TIPO DE CAMBIO
   Dim RsTipoCambio As Recordset
   
   Isql = "select * from TTIPOCAMBIO WHERE CONVERT(NVARCHAR,fFecha,103)= '" & FechaServidorTipoCambio() & "' "
   Set RsTipoCambio = Lib.OpenRecordset(Isql, Cn)
     
   If RsTipoCambio.RecordCount = 0 Then
                     
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If

                  oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, 1
                  oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, ""
                  oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, 0
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
   End If
End Sub

Private Sub ValidaHARDkey()
    On Error GoTo ErrHARDkey
    
    clave1 = "11111"
    clave2 = "22222"
    
    If lHARDkey Then
        '************Validacion Hard Key***************************
        Dim verif2 As Boolean
        Dim str As String
        Dim verif As String
     
        hk.SetClaves clave1, clave2
        verif = hk.IniciaConexion(Aplicacion.Adicion)
     
        If verif <> "" Then
            MsgBox verif, vbCritical, "Aviso"
            End
        End If
    
'        str = hk.ValidaFechaVencimiento
'
'        If str = "5" Then
'            MsgBox "No se encontro la llave", vbCritical, "Aviso"
'            End
'        End If
'
'        If Len(str) > 5 Then
'            MsgBox str, vbExclamation, "Aviso"
'        End If
'
'        If str = "4" Then
'            MsgBox "Su licencia ha caducado.", vbCritical, "Aviso"
'            verif2 = hk.FinalizarConexion(Aplicacion.Adicion)
'            End
'        End If
        '*********************************************************
    End If
ErrHARDkey:
End Sub


Private Sub FinalizaConexionHK()
    Dim Result As Boolean
    Result = hk.FinalizarConexion(Aplicacion.Adicion)
End Sub






