Attribute VB_Name = "modAuditoria"
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
    Dim sUserName As String
    Dim sPassword As String
    
    sUserName = "sa"
    sPassword = "sistemas"
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 300
   Cn.Open
   
   frmFlash.Label5.Caption = "Módulo de Auditoría"
   frmFlash.Show vbModal
    
   'Configuración
   'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Inforest;Data Source=LUIS
   'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inforest\Data\inforest.mdb;Persist Security Info=False
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "INFOREST"
   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!Impuesto1), 0, RsParametro!Impuesto1)
   nPorcentaje2 = IIf(IsNull(RsParametro!Impuesto2), 0, RsParametro!Impuesto2)
   nPorcentaje3 = IIf(IsNull(RsParametro!Impuesto3), 0, RsParametro!Impuesto3)
   nTiempo = IIf(IsNull(RsParametro!nTiempo), 0, RsParametro!nTiempo)
   nChkTiempo = IIf(IsNull(RsParametro!nChkTiempo), 0, RsParametro!nChkTiempo)
   nDelivery = IIf(IsNull(RsParametro!nDelivery), 0, RsParametro!nDelivery)
   nllevar = IIf(IsNull(RsParametro!nllevar), 0, RsParametro!nllevar)
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
   wObliga = IIf(IsNull(RsParametro!lObliga), False, RsParametro!lObliga)
   
   Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
   If RsTc.RecordCount <> 0 Then
      nTipoCambio = IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta)
   Else
      nTipoCambio = 0
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
   wInicio = True
   frmAcceso.Caption = "Inforest módulo de Auditoría v." & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
      mdiAuditoria.Show
   End If
End Sub

Public Sub Centrar(frmNombre As Form)
    With frmNombre
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
End Sub
