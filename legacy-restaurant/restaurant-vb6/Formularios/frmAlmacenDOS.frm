VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlmacenDOS 
   Caption         =   "Transferencia de Almacén"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmAlmacenDOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOpcion 
      Caption         =   " Rango de Fechas a Transferir "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   5265
      Begin VB.CheckBox chkPlato 
         Caption         =   "Actualización de Platos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   0
         Top             =   360
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1530
         TabIndex        =   3
         Top             =   1283
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24510465
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1530
         TabIndex        =   1
         Top             =   848
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24510465
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3510
         TabIndex        =   2
         Top             =   870
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   24510467
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3510
         TabIndex        =   4
         Top             =   1305
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   24510467
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   938
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1373
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Transferir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2340
      Picture         =   "frmAlmacenDOS.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3870
      Picture         =   "frmAlmacenDOS.frx":1B7C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmAlmacenDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCriterio As String
Dim sDBF2 As String

Private Sub chkPlato_Click()
    If chkPlato.Value = 1 Then
       dtpFecIni.Enabled = False
       dtpFecFin.Enabled = False
       dtpHorIni.Enabled = False
       dtpHorFin.Enabled = False
    Else
       dtpFecIni.Enabled = True
       dtpFecFin.Enabled = True
       dtpHorIni.Enabled = True
       dtpHorFin.Enabled = True
    End If
End Sub

Private Sub cmdOpcion_Click()
   If chkPlato.Value = 1 Then
      Dim CnPlato As Connection
      Dim sCodProd As String
      Set CnPlato = New ADODB.Connection
      CnPlato.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Plato DBF;Mode=ReadWrite;Initial Catalog=" & sDBF2
      CnPlato.Open
   
      'Transfiere La Tabla de Platos
      Dim RsBase As Recordset
      Dim RsCosto As Recordset
      Dim nCont As Integer
      Dim nTot As Integer
      Dim nGasto As Double
      Dim nManoObra As Double
      Dim nPrecioCosto As Double
      Dim nInsumo As Double
      
      Screen.MousePointer = vbHourglass
      Set RsBase = Lib.OpenRecordset("SELECT tCodigoProducto, tDetallado, tResumido, nPrecioCosto, nPrecioVenta, tDescargo, tEnlace, tMoneda, lImpuesto1, lImpuesto2 from TPRODUCTO", Cn)
      Set RsCosto = Lib.OpenRecordset("SELECT COD_AMARRE, PRE_COSTO, G_GENERAL, PRE_OPERA, PG_GENERAL, PPRE_OPERA from TSCOS01", CnAlmacen)
      
      If RsBase.RecordCount > 0 Then
         Cn.Execute "update TPRODUCTO set nGasto=0, nManoObra=0, nPrecioCosto=0, nInsumo=0"
         RsBase.MoveFirst
         nCont = 0
         nTot = RsBase.RecordCount
         Do While Not RsBase.EOF
            nCont = nCont + 1
            mdiAdministracion.StatusBar.Panels(1).Text = "Actualizando " & nCont & " de " & nTot
            sCodProd = "2" & Mid(RsBase!tCodigoProducto, 4, 4)
            RsCosto.Filter = "COD_AMARRE='" & sCodProd & "'"
            If RsCosto.RecordCount = 0 Then
               nGasto = 0
               nManoObra = 0
               nPrecioCosto = 0
            Else
               nGasto = RsCosto!G_GENERAL
               nManoObra = RsCosto!PRE_OPERA
               nInsumo = RsCosto!PRE_COSTO
               nPrecioCosto = nGasto + nManoObra + nInsumo
            End If
                        
            Cn.Execute "update TPRODUCTO set nGasto=" & nGasto & ", nManoObra=" & nManoObra & ", nPrecioCosto=" & nPrecioCosto & ", nInsumo =" & nInsumo & " where tCodigoProducto='" & RsBase!tCodigoProducto & "'"
            
            If Calcular("select COD_PROD as Codigo From TSHTA12 where COD_PROD='" & sCodProd & "'", CnPlato) <> "0" Then
               CnPlato.Execute "UPDATE TSHTA12 set DESC_COMP='" & UCase(RsBase!tDetallado) & "', DESC_RESU='" & UCase(RsBase!tResumido) & "', " & _
                               "PREC_COSTO=" & RsBase!nPrecioCosto & ",PREC_VENTA=" & RsBase!nPrecioVenta & ", RECETA='" & RsBase!tDescargo & "', COD_AMARRE='" & RsBase!tenlace & "', MONEDA='" & IIf(RsBase!tMoneda = "01", "S", "N") & "', " & _
                               "AFECTO_IGV='" & IIf(RsBase!lImpuesto1, "S", "N") & "', AFECTO_SER='" & IIf(RsBase!lImpuesto2, "S", "N") & "' " & _
                               "WHERE COD_PROD='" & sCodProd & "'"
            Else
               CnPlato.Execute "INSERT INTO TSHTA12 (COD_PROD, DESC_COMP, DESC_RESU, PREC_COSTO, PREC_VENTA, RECETA, COD_AMARRE, COD_TIPO, CODSEG, CODSUB, MONEDA, MEDIDA, AFECTO_IGV, AFECTO_SER, COD_PROD2, MONEDA, AFECTO_IGV, AFECTO_SER) " & _
                               "values ('" & sCodProd & "', '" & UCase(RsBase!tDetallado) & "', '" & UCase(RsBase!tResumido) & "', " & RsBase!nPrecioCosto & ", " & RsBase!nPrecioVenta & ", '" & RsBase!tDescargo & "', '" & RsBase!tenlace & "','','','','','','','','','" & IIf(RsBase!tMoneda = "01", "S", "N") & "', '" & IIf(RsBase!lImpuesto1, "S", "N") & "', '" & IIf(RsBase!lImpuesto2, "S", "N") & "')"
            End If
            RsBase.MoveNext
         Loop
         
      End If
      Set RsBase = Nothing
      Set RsCosto = Nothing
      Set CnPlato = Nothing
      Screen.MousePointer = vbDefault
   Else
      Dim RsTrans As Recordset
      Dim RsReceta As Recordset
      Dim RsRecDetalle As Recordset
      
      Screen.MousePointer = vbHourglass
      mdiAdministracion.StatusBar.Panels.Item(1).Text = "Iniciando Transferencia...."
      
      Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.vArea.tValor AS tCosto " & _
             "FROM dbo.TPRODUCTO LEFT OUTER JOIN dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
             "where (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and (isnull(lTransferido,'0') = '0' or ltransferido = 0 ) and tEstadoItem='N' and fFecha >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"
      Set RsTrans = Lib.OpenRecordset(Isql, Cn)
                                        
      If RsTrans.RecordCount > 0 Then
         RsTrans.MoveFirst
         Do While Not RsTrans.EOF
            mdiAdministracion.StatusBar.Panels.Item(1).Text = "Transfiriendo Movimiento " & RsTrans.AbsolutePosition & " de " & RsTrans.RecordCount
            
            Select Case RsTrans!tDescargo
                   Case Is = "D"
                        'Directo
                        Descargo RsTrans!tCosto, RsTrans!tCodigoPedido, RsTrans!tenlace, RsTrans!nCantidad, RsTrans!fFecha, "D"
                        
                   Case Is = "R"
                        'Receta
                        Set RsReceta = Lib.OpenRecordset("select * from TSCOS01 where COD_AMARRE ='2" & Mid(RsTrans!tCodigoProducto, 4, 4) & "'", CnAlmacen)
                        If RsReceta.RecordCount > 0 Then
                           Set RsRecDetalle = Lib.OpenRecordset("select * from TSCOS02 where COD_PLATO ='" & RsReceta!COD_PLATO & "'", CnAlmacen)
                           If RsRecDetalle.RecordCount > 0 Then
                              Do While Not RsRecDetalle.EOF
                                 Descargo RsReceta!COD_AREA, RsTrans!tCodigoPedido, RsRecDetalle!COD_PROD, RsRecDetalle!CAN_PROD * RsTrans!nCantidad, RsTrans!fFecha, "R"
                                 RsRecDetalle.MoveNext
                              Loop
                           End If
                        End If
                           
                   Case Else
                        'Sin descargo
            End Select
                                 
            Cn.Execute "update DPEDIDO set lTransferido = 1 where tCodigoPedido = '" & RsTrans!tCodigoPedido & "' and tItem = '" & RsTrans!tItem & "'"
            RsTrans.MoveNext
         Loop
         Screen.MousePointer = vbDefault
         MsgBox "Se realizó " & Trim(Str(RsTrans.RecordCount)) & " Transferencia(s) con Exito", vbInformation, "Transferencia"
         
      Else
         Screen.MousePointer = vbDefault
         MsgBox "No existe Ventas nuevas a transferir", vbExclamation, "Transferencia"
      End If
      Set RsTrans = Nothing
   End If
   
   mdiAdministracion.StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
End Sub

Private Sub cmdSalir_Click(Index As Integer)
   Unload Me
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Set CnAlmacen = New Connection
   CnAlmacen.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Almacen DBF;Mode=ReadWrite;Initial Catalog=" & sDBF
   CnAlmacen.Open
   
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   dtpFecIni.Enabled = False
   dtpFecFin.Enabled = False
   dtpHorIni.Enabled = False
   dtpHorFin.Enabled = False
End Sub

Public Sub Descargo(Area As String, Documento As String, Producto As String, Cantidad As Double, Fecha As Date, TipoDescargo As String)
   Dim RsStock As Recordset
   Dim RsKardex As Recordset
   Dim RsProducto As Recordset
   Dim wStock As String
   Dim wKardex As String
   Dim nStock As Double
      

      
   If IsNull(Area) Or Len(Area) < 1 Then
      Exit Sub
   End If
   
   Set RsProducto = Lib.OpenRecordset("select * from TSHAL02 where COD_PROD='" & Producto & "'", CnAlmacen)
   If RsProducto.RecordCount = 0 Then
      Exit Sub
   End If
   
   wStock = "TSS01" & Area
   wKardex = "TSK01" & Area
   
   Set RsKardex = Lib.OpenRecordset("select * from " & wStock & " where COD_PROD ='" & Producto & "'", CnAlmacen)
   
   If RsKardex.RecordCount = 0 Then
      nStock = 0
      CnAlmacen.Execute "Insert into " & wStock & ".DBF ( COD_PROD, STOCK_ACT ) Values ('" & Producto & "', - " & Round((Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!FACT_SAL)), 3) & ")"
   Else
      nStock = RsKardex!STOCK_ACT
      CnAlmacen.Execute "Update " & wStock & ".DBF set STOCK_ACT = STOCK_ACT - " & Round((Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!FACT_SAL)), 3) & " where COD_PROD ='" & Producto & "'"
   End If
   
   Isql = "Insert into " & wKardex & ".DBF " & _
          "(COD_PROD, TIPO_DOC, COD_AREA, DOCUMENTO, COSTO_UNIT, STOCK_ACT, ULT_STOCK, USUARIO, HORA, " & _
          "PRE_PROM, SALIDA, V_SALIDA, ESTADO, FECHA, INGRESO, V_INGRESO, TOT_SALDO) " & _
          "Values ( '" & Producto & "', 'C', " & _
                   "'" & Area & "', " & _
                   "'" & Documento & "', " & _
                         RsProducto!PREC_SAL & ", " & _
                         nStock - Round((Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!FACT_SAL)), 3) & ", " & _
                         nStock & ", " & _
                   "'" & sUsuario & "', " & _
                   "'" & Time() & "', " & _
                   RsProducto!PREC_SAL & ", " & _
                   Round((Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!FACT_SAL)), 3) & ", " & _
                   Round((Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!FACT_SAL)) * RsProducto!PREC_SAL, 3) & ", " & _
                   "'P', {" & Format(Fecha, "MM/dd/yyyy") & "}, 0, 0, 0 )"
 
   CnAlmacen.Execute Isql

   Set RsStock = Nothing
   Set RsKardex = Nothing
   Set RsProducto = Nothing
End Sub


