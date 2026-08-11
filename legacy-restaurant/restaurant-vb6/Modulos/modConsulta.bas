Attribute VB_Name = "modConsulta"
Public hk As New License
 
 Sub Main()
 Dim existe As Boolean
 Dim AnoMesSv As String
 Dim directorioSistema   As String
 Dim archI As String
    lVersionEducativa = False
    
    lHARDkey = False
    
    moduloUso = "Consulta"

    ultimoConectado = True
    
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    
    
    'FACTURACION_E_PERU
    sRutaFE = Trim(LeerIni(App.path + "\FACTURACION.INI", "Conexion", "SERVIDOR", "."))
    sMDBFE = Trim(LeerIni(App.path + "\FACTURACION.INI", "Conexion", "BASEDATO", "BDEFACT"))
    '-----------------
   
    
    Dim RsParametro As Recordset
    Dim RsTc As Recordset
    Dim RsCaja As Recordset
    Dim RsBusca As Recordset
    Dim RsFactura As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    'auditoria
    nCorrelativoAcceso = 0
    tModuloSeg = "14" 'XX= CODIGO DEL MODULO DE LA TABLA DE MMODULO DE SEGURIDAD
    Set CnSeg = New Connection
      
    CnSeg.Provider = "SQLOLEDB"
    CnSeg.CursorLocation = adUseServer
    CnSeg.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=INFSEGURIDAD"
    CnSeg.CommandTimeout = 250
    
    'auditoria

    
    
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
    Cn.CommandTimeout = 300
    Cn.Open
      
        
     'dic 2010 un exe varias bd
    Set cnDefault = Cn
    'dic 2010 un exe varias bd
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
    mstrFile = App.path + "\" & App.EXEName & ".exe"
    
    Dim sActRuta As String
    sActRuta = Trim(LeerIni(App.path + "\INFOREST.INI", "Actualizador", "RUTA", ".")) & "\" & App.EXEName & ".exe"

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
        Call IniciarActualizador("CONSULTAS")
        End
    
'        'CREA LA CARPETA PARA BACKUP DE EXES
'        Dim Backup As String
'        Backup = App.path + "\ExesHistoricos"
'
'        'CREA LA CARPETA Y VALIDA QUE NO EXISTA
'        If Dir(Backup, vbDirectory) = "" Then
'           MkDir (Backup)
'        End If
'
'        Shell App.path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
        End
    Else
         'Actualizador incrustado ELDCQ
        Call IniciarActualizador("INICIO")
        'Sleep (3000)
    End If
       
    '------
    '=================================================
    pais = ObtienePais
    '=================================================
    
    'Diccionario
    
    Mesero = DiccionarioValor("Mesero")
    
    '===================================================

    'Configuración
   'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Inforest;Data Source=LUIS
   'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inforest\Data\inforest.mdb;Persist Security Info=False
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "CONSULTA"
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
   sPiePreCuenta = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sFax = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
   sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   sMail = IIf(IsNull(RsParametro!tEmail), "", RsParametro!tEmail)
   nCabecera = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
   nDetalle = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
   sRUC = SoloNumeros(IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria))
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
   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
   sClub = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   lCierre = IIf(IsNull(RsParametro!lCierre), False, RsParametro!lCierre)
   
   
   
   lCambiarContrasena = Calcular("select lCambioContrasena as Codigo from TPARAMETRO", Cn)
   
   locultaimpuesto2 = IIf(IsNull(RsParametro!lvisibleimpuesto2), 0, RsParametro!lvisibleimpuesto2)
   locultaimpuesto3 = IIf(IsNull(RsParametro!lvisibleimpuesto3), 0, RsParametro!lvisibleimpuesto3)
   
       'HUELLA
     lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
      lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)
    'KDS
   'FACTURACION ELECTRONICA
   lFacturacionE = IIf(IsNull(RsParametro!lFacturacionE), False, RsParametro!lFacturacionE)
   tCodigoFE = IIf(IsNull(RsParametro!tCodigoFE), "000", RsParametro!tCodigoFE)
   tPieDocumento1 = IIf(IsNull(RsParametro!tPieDocumento1), " ", RsParametro!tPieDocumento1)
   lAmbienteProduccion = IIf(IsNull(RsParametro!lAmbienteFE), False, RsParametro!lAmbienteFE)
   
   'agenteretencion
   tTextoAgenteRetencion = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   
   lImpresionCodigoBarras = IIf(IsNull(RsParametro!lImprimeCodigoBarras), False, RsParametro!lImprimeCodigoBarras)
   
   lNcOfisis = IIf(IsNull(RsParametro!lNcOfisis), False, RsParametro!lNcOfisis)
   
   'TCANALVENTA
   sBoton1 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
   sBoton2 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
   sBoton3 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
   sBoton4 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
   sBoton5 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)
   nDias = IIf(IsNull(RsParametro!nDias), 2, RsParametro!nDias)
   nPunto = IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto)
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   
   'club
   lClub = IIf(IsNull(RsParametro!lClub), False, RsParametro!lClub)
   obtieneVencimientoConexiones
   
   'fe
    lQRFE = IIf(IsNull(RsParametro!lCodigoQrFE), 0, IIf(RsParametro!lCodigoQrFE = True, 1, 0))
    RutaImgFE = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE)
   
   'fepaperllees
   lFEpape = IIf(IsNull(RsParametro!lFEpape), 0, IIf(RsParametro!lFEpape = True, 1, 0))
   lDesactivaNCFP = IIf(IsNull(RsParametro!lDesactivaNCFP), 0, IIf(RsParametro!lDesactivaNCFP = True, 1, 0))
   sLey1 = IIf(IsNull(RsParametro!tLey1), "", RsParametro!tLey1)
   sValorLey1 = IIf(IsNull(RsParametro!tValorLey1), "", RsParametro!tValorLey1)
   'FACTURACION_E_PERU
   On Error GoTo FacturacionIni
   
   If lFacturacionE And lFEOfisis = False And lFEpape = False Then
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
    If Err.Number = "-2147467259" Then
       MsgBox "Facturacion Electronica, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    End If
    '-------------------------------
    
   
   
   
   
   'KDS Para reporte KDS
   lKDS = IIf(IsNull(RsParametro!lKDS), False, RsParametro!lKDS)
   If (lKDS) And sModulo <> "CONSULTA" Then
        sBump = IIf(IsNull(RsParametro!tBump), "", RsParametro!tBump)
        KDS_ListarBumpNotification (sBump)
   End If
      
   Set CnAlmacen = New Connection
   Set CnAlmacenRemoto = New Connection
   Set cnAlmacenDefault = New Connection
   
   If lAlmacen Then
      ' Configuracion ini
      sAlmacenRuta = Trim(LeerIni(App.path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
      sAlmacenMDB = Trim(LeerIni(App.path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
      sLocal = Trim(LeerIni(App.path + "\ALMACEN.INI", "Configuracion", "LOCAL", "01"))
      
      On Error GoTo ValidaAlmacen
        'GCAA 15072021 - CASO CORP.FERNANDEZ
      BDLink = Trim(LeerIni(App.path + "\ALMACEN.INI", "BDLINK", "ACTIVO", "0"))
      CnAlmacen.Provider = "SQLOLEDB"
      CnAlmacen.CursorLocation = adUseServer
      CnAlmacen.ConnectionString = "User ID=" & sUserName & _
                                   ";password=" & sUserPassword & _
                                   ";Data Source=" & sAlmacenRuta & _
                                   ";Initial Catalog=" & sAlmacenMDB
      CnAlmacen.Open
      Set cnAlmacenDefault = CnAlmacen
        localConectado = Calcular("select isnull(tresumido,'0') as codigo from vlocalidades where ip='" & sRuta & "' and bdinf='" & sMDB & "'", CnAlmacen)
      
'      '''''''''''''''''''''''''' almacen remoto
'      Dim verificaAlmacenRemoto As String
'      verificaAdministracionCentralizada = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "REMOTO", "OFF"))
'      If verificaAdministracionCentralizada = "ON" Then
'            lAlmacenRemoto = True
'            sRutaAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "SERVIDOR", "LOCAL"))
'            sMDBAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "BASEDATO", "ALMACEN"))
'      End If
      ''''''''''''''''''''''''''
ValidaAlmacen:
        If Err.Number = "-2147467259" Then
           'MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
            If MsgBox("Almacen, SQL Server no Encontrado, Falla de Conectividad!!! " & vbNewLine & " ¿Desea Continuar sin enlace a Almacen?" & vbNewLine & " de continuar con el proceso, se informa que los costos de los productos " & vbNewLine & " vendidos sin enlace a almacen no se grabaran!!! ", vbYesNo, "Almacen") = vbNo Then
                Exit Sub
            Else
                Call Log_Inforest("Almacen", "CONEXIONALMACEN", TipoIni, "Caja: " + sCaja, Environ("computername"), "Ingreso sin conexion de almacen a Inforest", "", "FALLA PROCESO", "INICIO")
                lAlmacen = False
            End If
        End If
        
   End If
    ' UN EXE VARIAS BD
    multiLocal = IIf(IsNull(RsParametro!lmultilocal), False, RsParametro!lmultilocal)
    'dic 2010 un exe varias bd

    If localConectado = "0" Or localConectado = "" Then
            localConectado = sRazonComercial
    End If
    
    Call ValidaHARDkey 'OO
    'frmFlash.Label5.Caption = "Módulo de Consultas"
    'frmFlash.Show vbModal
                               
                             
   'Proceso de Caja
   Set RsCaja = Lib.OpenRecordset("select * from TCAJA", Cn)
   If RsCaja.RecordCount <> 0 Then
      RsCaja.MoveFirst
      RsCaja.Find ("tCaja='" & sCaja & "'")
      If RsCaja.EOF Then
         MsgBox "Error Fatal: No existe Caja Configurada", vbCritical, sMensaje
         End
      Else
         sPreCuenta = IIf(IsNull(RsCaja!tPrecuenta), "001", RsCaja!tPrecuenta)
         
         'descripcion alternativa
         lDescripcionAlternativa = IIf(IsNull(RsCaja!lActivaImpDscAlternativa), False, RsCaja!lActivaImpDscAlternativa)
        lImprimeImagCabPrecuenta = IIf(IsNull(RsCaja!lImprimeImagCabPrecuenta), False, RsCaja!lImprimeImagCabPrecuenta)
        lImprimeImagPiePrecuenta = IIf(IsNull(RsCaja!lImprimeImagPiePrecuenta), False, RsCaja!lImprimeImagPiePrecuenta)
        lBuscarPedidoBarraMesero = IIf(IsNull(RsCaja!lBuscarpedidoBarra), False, RsCaja!lBuscarpedidoBarra)
        lActivaVoucher = IIf(IsNull(RsCaja!lActivaVoucher), 0, RsCaja!lActivaVoucher)
      End If
   Else
      MsgBox "Error Fatal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   Set RsTc = Nothing
   Set RsCaja = Nothing
   Set RsParametro = Nothing
   Set RsBusca = Nothing
   Set RsFactura = Nothing
   wInicio = False
   frmAcceso.Caption = "Inforest módulo de Consultas v." & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
      mdiConsulta.Show
   End If
   Exit Sub
InforestIni:
    If Err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox Err.Description & ":" & Err.Number
    End If
    Exit Sub
AlmacenIni:
    If Err.Number = "-2147467259" Then
       MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    Else
       MsgBox Err.Description & ":" & Err.Number
    End If
    Exit Sub
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
        verif = hk.IniciaConexion(Aplicacion.Consultas)
            
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
'            verif2 = hk.FinalizarConexion(Aplicacion.Consultas)
'            End
'        End If
        '*********************************************************
    End If
ErrHARDkey:
End Sub

Private Sub FinalizaConexionHK()
    Dim Result As Boolean
    Result = hk.FinalizarConexion(Aplicacion.Consultas)
End Sub
