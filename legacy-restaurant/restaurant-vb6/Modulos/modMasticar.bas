Attribute VB_Name = "modMasticar"
Public fFechaI As Date
Public fFechaF As Date
Public nMonto As Double
Public lCopia As Boolean
Public nSuma As Double

Sub Main()
    'lIf App.PrevInstance = True Then
    '   MsgBox "El Programa se encuentra en ejecución", vbExclamation, "Mensaje Inforest"
    '   End
    'End If

    ' Configuracion ini
    Open "INFOREST.INI" For Input As #1                ' Abre el archivo para recibir los datos.
    Do While Not EOF(1)                                ' Repite el bucle hasta el final del archivo.
       Input #1, sRuta, sMDB, sCaja, sSalon, sEmpresa  ' Lee el carácter en dos variables
    Loop
    Close #1   ' Cierra el archivo.
            
    sRuta = IIf(Right(Trim(sRuta), 1) = "\", sRuta, sRuta)
    sMDB = IIf(UCase(Right(Trim(sMDB), 4)) = ".MDB", sMDB, sMDB)
            
    Dim RsParametro As Recordset
    Dim RsEnvio As Recordset
        
    Set Lib = New Libreria16.Applications
    Dim sUserName As String
    Dim sPassword As String
    
    sUserName = "sa"
    sPassword = "sistemas"
    
    Set Cn = New ADODB.Connection
    Cn.Provider = "SQLOLEDB"
    'Cn.CursorLocation = adUseServer

    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
    Cn.CommandTimeout = 300
    Cn.Mode = adModeShareExclusive
    Cn.IsolationLevel = adXactIsolated
    
    Cn.Open
    
    Set CnDBF = New Connection
    CnDBF.Provider = "SQLOLEDB"
    CnDBF.CursorLocation = adUseServer
    CnDBF.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=COPIA"
    CnDBF.CommandTimeout = 300
    CnDBF.Open
    
   'Configuración
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   nPorcentaje1 = IIf(IsNull(RsParametro!Impuesto1), 0, RsParametro!Impuesto1)
   nPorcentaje2 = IIf(IsNull(RsParametro!Impuesto2), 0, RsParametro!Impuesto2)
   nPorcentaje3 = IIf(IsNull(RsParametro!Impuesto3), 0, RsParametro!Impuesto3)
   sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tPassword), "", RsParametro!tPassword)))
   
   sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sPie = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   
   Set RsEnvio = Lib.OpenRecordset("select * from MENVIO order by fFinal", Cn)
   
   If RsEnvio.EOF Then
      MsgBox "Error: No existe fecha de Envio", vbCritical, sMensaje
      End
   End If
   RsEnvio.MoveLast
   
   If RsEnvio!lCierre = True Then
      frmFecha.dtpFecha1.Value = RsEnvio!fFinal + 1
      frmFecha.Show vbModal
   Else
      fFechaI = RsEnvio!finicio
      fFechaF = RsEnvio!fFinal
      nMonto = IIf(IsNull(RsEnvio!nMonto), 0, RsEnvio!nMonto)
      lCopia = IIf(IsNull(RsEnvio!lCopia), False, RsEnvio!lCopia)
   End If
   
   Set RsParametro = Nothing
   Set RsEnvio = Nothing
   sUsuario = "MASTER"
   mdiMasticar.Show
End Sub


Public Sub EliminaTemporal()
   On Error Resume Next
   Screen.MousePointer = vbHourglass
   mdiMasticar.StatusBar.Panels(1).Text = "Eleminando Temporales..."
   Dim RsTemp As Recordset
   Isql = "SELECT NAME AS Nombre FROM SYSOBJECTS WHERE TYPE='U' AND NAME LIKE 'TMP%'"
   Set RsTemp = Lib.OpenRecordset(Isql, Cn)
   
   If RsTemp.RecordCount > 0 Then
      Do While Not RsTemp.EOF
         Cn.Execute "Drop Table " & RsTemp!Nombre
         RsTemp.MoveNext
      Loop
   End If
   Set RsTemp = Nothing
   Cn.Execute "DBCC SHRINKDATABASE (" & sMDB & ", 0)"
   mdiMasticar.StatusBar.Panels(1).Text = "Sistema Listo"
   Screen.MousePointer = vbDefault
End Sub
