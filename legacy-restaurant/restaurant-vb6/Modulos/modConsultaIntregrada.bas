Attribute VB_Name = "modConsultaIntregrada"
Public hk As New License
 
 Sub Main()
  Dim existe As Boolean
 Dim AnoMesSv As String
 Dim directorioSistema  As String
 Dim archI As String
    lHARDkey = False
    lLeeCsg = False
    
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    
    Dim RsParametro As Recordset
    Dim RsTc As Recordset
    Dim RsCaja As Recordset
    Dim RsBusca As Recordset
    Dim RsFactura As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "infhotel"
    sUserPassword = "infh2566"
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
    Cn.CommandTimeout = 300
    Cn.Open
        
   
     '====================================================
    If lLeeCsg = True Then
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
    '===================================================
    
    
    'Configuración
   'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Inforest;Data Source=LUIS
   'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inforest\Data\inforest.mdb;Persist Security Info=False
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "INTEGRACION"
   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!Impuesto1), 0, RsParametro!Impuesto1)
   nPorcentaje2 = IIf(IsNull(RsParametro!Impuesto2), 0, RsParametro!Impuesto2)
   nPorcentaje3 = IIf(IsNull(RsParametro!Impuesto3), 0, RsParametro!Impuesto3)
   nTiempo = IIf(IsNull(RsParametro!nTiempo), 0, RsParametro!nTiempo)
   nChkTiempo = IIf(IsNull(RsParametro!nChkTiempo), 0, RsParametro!nChkTiempo)
   nDELIVERY = IIf(IsNull(RsParametro!nDELIVERY), 0, RsParametro!nDELIVERY)
   nLlevar = IIf(IsNull(RsParametro!nLlevar), 0, RsParametro!nLlevar)
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sPie = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   sPiePreCuenta = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   sMail = IIf(IsNull(RsParametro!tEmail), "", RsParametro!tEmail)
   sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tPassword), "", RsParametro!tPassword)))
   sElimina = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
   sClub = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   lCierre = IIf(IsNull(RsParametro!lCierre), False, RsParametro!lCierre)
   sBoton1 = IIf(IsNull(RsParametro!tBoton1), "", RsParametro!tBoton1)
   sBoton2 = IIf(IsNull(RsParametro!tBoton2), "", RsParametro!tBoton2)
   sBoton3 = IIf(IsNull(RsParametro!tBoton3), "", RsParametro!tBoton3)
   sBoton4 = IIf(IsNull(RsParametro!tBoton4), "", RsParametro!tBoton4)
   sBoton5 = IIf(IsNull(RsParametro!tBoton5), "", RsParametro!tBoton5)
   nDias = IIf(IsNull(RsParametro!nDias), 2, RsParametro!nDias)
   nPunto = IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto)
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   
    'Coneccion para Almacen
    If lAlmacen Then
        xRuta = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
        xMDB = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
    
        Set CnAlmacen = New Connection
        CnAlmacen.Provider = "SQLOLEDB"
        CnAlmacen.CursorLocation = adUseServer
        CnAlmacen.ConnectionString = "User ID=" & sUserName & _
                                     ";password=" & sUserPassword & _
                                     ";Data Source=" & xRuta & _
                                     ";Initial Catalog=" & xMDB
        CnAlmacen.CommandTimeout = 300
        CnAlmacen.Open
    End If
    
    lAlmacen = False
    Call ValidaHARDkey 'OO
    frmFlash.Label5.Caption = "Módulo de Consultas Integradas"
    frmFlash.Show vbModal
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
      End If
   Else
      MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   Set RsTc = Nothing
   Set RsCaja = Nothing
   Set RsParametro = Nothing
   Set RsBusca = Nothing
   Set RsFactura = Nothing
    
   wInicio = False
   frmAcceso.Caption = "Inforest Módulo de Consultas Integradas v." & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
       frmServidores.Show vbModal
         If wEnter = True Then
               mdiConsultaIntegrada.Show
          End If
   End If
   Exit Sub
InforestIni:
    If err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    If lHARDkey Then
       FinalizaConexionHK
    End If
    Exit Sub
AlmacenIni:
    If err.Number = "-2147467259" Then
       MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    If lHARDkey Then
       FinalizaConexionHK
    End If
    Exit Sub
End Sub

Private Sub FinalizaConexionHK()
    Dim result As Boolean
    result = hk.FinalizarConexion(Aplicacion.Consultas)
End Sub

Private Sub ValidaHARDkey()
    On Error GoTo ErrHARDkey
    clave1 = "67332"
    clave2 = "5877"
    
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
            
        str = hk.ValidaFechaVencimiento
            
        If str = "5" Then
            MsgBox "No se encontro la llave", vbCritical, "Aviso"
            End
        End If
            
        If Len(str) > 5 Then
            MsgBox str, vbExclamation, "Aviso"
        End If
            
        If str = "4" Then
            MsgBox "Su licencia ha caducado.", vbCritical, "Aviso"
            verif2 = hk.FinalizarConexion(Aplicacion.Consultas)
            End
        End If
        '*********************************************************
    End If
ErrHARDkey:
End Sub
