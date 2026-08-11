Attribute VB_Name = "modCheffControl"
 Sub Main()
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
        
    Dim RsParametro As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "Infhotel"
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
          
   frmFlash.Label5.Caption = "Módulo de Cheff Control"
   frmFlash.Show vbModal
          
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)

   Set RsParametro = Nothing
   
   wInicio = True
   Screen.MousePointer = vbDefault
   frmCheffControl.Show vbModal
End Sub
