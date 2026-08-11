Attribute VB_Name = "modPvCorp"
 Sub Main()
    ' Configuracion ini
    Open App.Path & "\INFOREST.INI" For Input As #1   ' Abre el archivo para recibir los datos.
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
       Input #1, sRuta, sMDB, sCaja, sSalon, sEmpresa   ' Lee el carácter en dos variables
    Loop
    Close #1   ' Cierra el archivo.
    sRuta = IIf(Right(Trim(sRuta), 1) = "\", sRuta, sRuta)
    sMDB = IIf(UCase(Right(Trim(sMDB), 4)) = ".MDB", sMDB, sMDB)
    
    Dim RsParametro As Recordset
    Dim RsTc As Recordset
    Dim RsCaja As Recordset
    Dim RsBusca As Recordset
    Dim RsFactura As Recordset
    
    Set Lib = New Libreria16.Applications
    
    'Configuración
    Dim sUserName As String
    Dim sPassword As String
    
    sUserName = "infhotel"
    sPassword = "infh2566"
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 300
   Cn.Open

   frmFlash.Label5.Caption = "Módulo Venta Corporativo"
   frmFlash.Show vbModal

   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
            
   sModulo = "PVCORP"
   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!Impuesto1), 0, RsParametro!Impuesto1)
   nPorcentaje2 = IIf(IsNull(RsParametro!Impuesto2), 0, RsParametro!Impuesto2)
   nPorcentaje3 = IIf(IsNull(RsParametro!Impuesto3), 0, RsParametro!Impuesto3)
   nTiempo = IIf(IsNull(RsParametro!nTiempo), 0, RsParametro!nTiempo)
   nChkTiempo = IIf(IsNull(RsParametro!nChkTiempo), 0, RsParametro!nChkTiempo)
   nDelivery = IIf(IsNull(RsParametro!nDelivery), 0, RsParametro!nDelivery)
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
   nFItem = IIf(IsNull(RsParametro!nItem), 0, RsParametro!nItem)
   nCabecera = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
   nDetalle = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
   lPrinter = IIf(IsNull(RsParametro!lPrinter), False, RsParametro!lPrinter)
   lLongitud = IIf(IsNull(RsParametro!lLongitud), False, RsParametro!lLongitud)
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   lRapido = IIf(IsNull(RsParametro!lRapido), False, RsParametro!lRapido)
   sBoton1 = IIf(IsNull(RsParametro!tBoton1), "", RsParametro!tBoton1)
   sBoton2 = IIf(IsNull(RsParametro!tBoton2), "", RsParametro!tBoton2)
   sBoton3 = IIf(IsNull(RsParametro!tBoton3), "", RsParametro!tBoton3)
   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   sClub = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   nPunto = IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto)
   lCierre = IIf(IsNull(RsParametro!lCierre), False, RsParametro!lCierre)
   nDecimal = IIf(IsNull(RsParametro!nDecimal), 2, RsParametro!nDecimal)
   nDias = IIf(IsNull(RsParametro!nDias), 2, RsParametro!nDias)
   lEquivalencia = IIf(IsNull(RsParametro!lEquivalencia), False, RsParametro!lEquivalencia)
   nTipoCambio = 0
   
    
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
         sGrupoDefault = IIf(IsNull(RsCaja!tGrupo), "01", RsCaja!tGrupo)
         lConsumo1 = IIf(IsNull(RsCaja!lConsumo1), False, RsCaja!lConsumo1)
         lConsumo2 = IIf(IsNull(RsCaja!lConsumo2), False, RsCaja!lConsumo2)
         lConsumo3 = IIf(IsNull(RsCaja!lConsumo3), False, RsCaja!lConsumo3)
      End If
   Else
      MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   'Proceso de Correlativo
   Isql = "select * from vTipoDocumento where Descripcion='FACTURA'"
   Set RsBusca = Lib.OpenRecordset(Isql, Cn)
   If RsBusca.RecordCount > 0 Then
      Isql = "select * from TTIPODOCUMENTOIMPRESORA where tCaja ='" & sCaja & "' and tTipoEmision='" & RsBusca!Codigo & "'"
      Set RsFactura = Lib.OpenRecordset(Isql, Cn)
      If RsFactura.RecordCount > 0 Then
         nFactura = RsFactura!tUltimoNumero
      Else
         nFactura = "Sin Correlativo"
      End If
   Else
     nFactura = "Sin Correlativo"
   End If
   
   Set RsTc = Nothing
   Set RsCaja = Nothing
   Set RsParametro = Nothing
   Set RsBusca = Nothing
   Set RsFactura = Nothing
   
   sTurno = "CORPORATIV"
   wInicio = True
   
   frmAcceso.Caption = "Inforest Módulo Venta Corporativo " & "v." & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
      mdiPvCorp.Show
   End If
End Sub
