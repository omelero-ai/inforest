Attribute VB_Name = "modDespachador"
 Sub Main()
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    sRutaCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "SERVIDOR", "."))
    sMDBCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "BASEDATO", "INFOREST"))
    
    sRutaWAP = Trim(LeerIni(App.Path + "\INFOREST.INI", "WebMobile", "SERVIDOR", "."))
    sMDBWAP = Trim(LeerIni(App.Path + "\INFOREST.INI", "WebMobile", "BASEDATO", "WEBAPP"))
    
    Dim RsParametro As Recordset
    Set Lib = New Libreria16.Applications
    sUserName = "Infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 300
   Cn.Open
          
          
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
    
    'auditoria
          
   frmFlash.Label5.Caption = "Módulo de Despachador"
   frmFlash.Show vbModal
          
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   'HUELLA
   lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
   lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)
   'KDS
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   tTarifaActualMotorizado = IIf(IsNull(RsParametro!tTarifaActualMotorizado), "", RsParametro!tTarifaActualMotorizado)
        '===================================================================
        obtieneVencimientoConexiones
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
            
            Shell App.Path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
            End
        End If
           

        validacionLicenciasInfhotel
    
   Set RsParametro = Nothing
        
        
   CD = Calcular("select lCD as Codigo from TCAJA where tCaja='" & sCaja & "'", Cn)
   
   WEBAP = Calcular("select ISNULL(lWebAp,0) as Codigo from TCAJA where tCaja='" & sCaja & "'", Cn)
   
   wInicio = True
   'frmAcceso.Caption = "Módulo del Despachador " & App.Major & "." & App.Minor & "." & App.Revision
   'frmAcceso.Show vbModal
   'If wEnter = True Then
   Screen.MousePointer = vbDefault
   'End If
'End Sub

   Set RsCaja = Lib.OpenRecordset("select * from TCAJA", Cn)
   If RsCaja.RecordCount <> 0 Then
      RsCaja.MoveFirst
      RsCaja.Find ("tCaja='" & sCaja & "'")
      If RsCaja.EOF Then
         MsgBox "Error Faltal: No existe Caja Configurada", vbCritical, sMensaje
         End
      Else
         sDespachador = IIf(IsNull(RsCaja!tDespachador), "001", RsCaja!tDespachador)
         lOrdenesRappi = IIf(IsNull(RsCaja!lOrdenesRappi), False, RsCaja!lOrdenesRappi)
      End If
   Else
      MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   mdiDespachador.Show
End Sub
