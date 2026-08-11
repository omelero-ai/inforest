Attribute VB_Name = "modGuias"
 Sub Main()
    ' Configuracion ini
    Open App.Path & "\INFOREST.INI" For Input As #1    ' Abre el archivo para recibir los datos.
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
       Input #1, sRuta, sMDB, sCaja, sSalon, sEmpresa   ' Lee el carácter en dos variables
    Loop
    Close #1   ' Cierra el archivo.
    sRuta = IIf(Right(Trim(sRuta), 1) = "\", sRuta, sRuta)
    sMDB = IIf(UCase(Right(Trim(sMDB), 4)) = ".MDB", sMDB, sMDB)
        
    Dim RsParametro As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "sa"
    sUserPassword = "sistemas"
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 300
   Cn.Open
          
   frmFlash.Label5.Caption = "Módulo de Guías de Remisión"
   frmFlash.Show vbModal
          
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!Impuesto1), 0, RsParametro!Impuesto1)
   nPorcentaje2 = IIf(IsNull(RsParametro!Impuesto2), 0, RsParametro!Impuesto2)
   nPorcentaje3 = IIf(IsNull(RsParametro!Impuesto3), 0, RsParametro!Impuesto3)
   
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   lLongitud = IIf(IsNull(RsParametro!lLongitud), False, RsParametro!lLongitud)
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   nFItem = IIf(IsNull(RsParametro!nItem), 11, RsParametro!nItem)

   Set RsParametro = Nothing
   
   wInicio = True
   frmAcceso.Caption = "Módulo de Guías de Remisión " & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
      mdiGuias.Show
   End If
End Sub
