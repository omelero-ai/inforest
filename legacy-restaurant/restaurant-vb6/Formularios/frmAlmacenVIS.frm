VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAlmacenVIS 
   BackColor       =   &H80000004&
   Caption         =   "Descargos Almacén"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frmAlmacenVIS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOpcion 
      Caption         =   " Rango de Fechas a Descargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   7185
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1530
         TabIndex        =   2
         Top             =   810
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
         Format          =   51773441
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1530
         TabIndex        =   0
         Top             =   375
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
         Format          =   51773441
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3510
         TabIndex        =   1
         Top             =   390
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
         Format          =   51773443
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3510
         TabIndex        =   3
         Top             =   825
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
         Format          =   51773443
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
         TabIndex        =   8
         Top             =   465
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
         TabIndex        =   7
         Top             =   900
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Descargar"
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
      Left            =   4230
      Picture         =   "frmAlmacenVIS.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1605
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
      Left            =   5760
      Picture         =   "frmAlmacenVIS.frx":1B7C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1605
      Width           =   1455
   End
End
Attribute VB_Name = "frmAlmacenVIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCriterio As String

Dim RsTransferencia As Recordset
Dim RsRecetaVenta As Recordset
Dim RsRecetaVentaDetalle As Recordset

Dim RsComboDetalle As Recordset

Dim RsProducto As Recordset
Dim RsSubStock As Recordset

Dim RsRecBaseBas As Recordset

Dim xTotal As Double

Private Sub cmdOpcion_Click()
   Screen.MousePointer = vbHourglass
   mdiAdministracion.StatusBar.Panels.Item(1).Text = "Iniciando Descargo..."
   
   xTotal = 0
   
   'Ventas a descargar
   Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.vArea.tValor AS tCosto, dbo.TPRODUCTO.lCombinacion, dbo.DPEDIDO.tItem " & _
          "FROM dbo.TPRODUCTO LEFT OUTER JOIN dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          "where (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and (isnull(lTransferido,'0') = '0' or ltransferido = 0 ) and tEstadoItem='N' and fFecha >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"
   Set RsTransferencia = Lib.OpenRecordset(Isql, Cn)
   
   Set RsRecetaVenta = Lib.OpenRecordset("select * from MRECETAVENTA", CnAlmacen)
   Set RsRecetaVentaDetalle = Lib.OpenRecordset("select * from vRecetaVentaDetalle where lDescargo=1", CnAlmacen)
      
   If RsTransferencia.RecordCount > 0 Then
      Set RsProducto = Lib.OpenRecordset("select * from TPRODUCTO", CnAlmacen)
      Set RsSubStock = Lib.OpenRecordset("select * from TSUBSTOCK", CnAlmacen)
      RsTransferencia.MoveFirst
      xTotal = RsTransferencia.RecordCount
      
      Do While Not RsTransferencia.EOF
         mdiAdministracion.StatusBar.Panels.Item(1).Text = "Transfiriendo Movimiento " & RsTransferencia.AbsolutePosition & " de " & xTotal
         
         'Procesos
         Select Case RsTransferencia!tDescargo
                Case Is = "D"  'Descargo Directo
                     Descargo RsTransferencia!tCosto, RsTransferencia!tCodigoPedido, RsTransferencia!tEnlace, RsTransferencia!nCantidad, "D"
                     
                Case Is = "R"  'Receta
                     DescargoReceta RsTransferencia!tEnlace, RsTransferencia!nCantidad
                  
                Case Is = "M"  'Menu
                     DescargoCombo
         End Select
                  
         Cn.Execute "update DPEDIDO set lTransferido = 1 where tCodigoPedido = '" & RsTransferencia!tCodigoPedido & "' and tItem = '" & RsTransferencia!tItem & "'"
         RsTransferencia.MoveNext
      Loop
      
      Screen.MousePointer = vbDefault
      MsgBox "Se realizó " & Trim(Str(RsTransferencia.RecordCount)) & " Transferencia(s) con Exito", vbInformation, "Transferencia"
   Else
      Screen.MousePointer = vbDefault
      MsgBox "No existe Ventas nuevas a transferir", vbExclamation, "Transferencia"
   End If
        
   Set RsTransferencia = Nothing
   Set RsRecetaVenta = Nothing
   Set RsRecetaVentaDetalle = Nothing
   
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
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   dtpFecIni.Enabled = True
   dtpFecFin.Enabled = True
   dtpHorIni.Enabled = True
   dtpHorFin.Enabled = True
End Sub

Public Sub Descargo(Area As String, Documento As String, Producto As String, Cantidad As Double, TipoDescargo As String)
   'Insumo de Descargo D=Directo / R=Receta
   
   Dim nStock As Double
   Dim nCorrela As Double
      
   If IsNull(Area) Or Len(Area) < 1 Then
      Exit Sub
   End If
      
   RsProducto.Filter = "tCodigoProducto='" & Producto & "'"
   If RsProducto.RecordCount = 0 Then
      Exit Sub
   End If
   
   'Descargo hacia el Almacen Central
   If Area = "000" Then
      nStock = RsProducto!nStockActual
      CnAlmacen.Execute "Update TPRODUCTO set nStockActual = nStockActual - " & (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", ncorrelativo = ncorrelativo + 1 where tCodigoProducto ='" & Producto & "'"

      nCorrela = Calcular("select max(nCorrelativo) as Codigo from MKARDEX where tCodigoProducto='" & Producto & "'", CnAlmacen)
      Isql = "Insert into MKARDEX " & _
             "(tCodigoArea, tCodigoProducto, nCorrelativo, fRegistro, tUsuario, tTipoDocumento, " & _
             "tDocumento, nIngreso, nSalida, nValor, nStockActual, nStockUltimo, tEstadoKardex, nPrecioPromedio, nPrecioCosto) " & _
             "Values ( '" & Area & "', " & _
                      "'" & Producto & "', " & _
                            nCorrela + 1 & ", " & _
                      " getdate(), " & _
                      " '" & sUsuario & "', " & _
                      "'95', " & _
                      "'" & Documento & "', 0, " & _
                            (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            RsProducto!npreciopromedio * (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            nStock - (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            nStock & ", " & _
                      "'01', " & RsProducto!npreciopromedio & ", " & RsProducto!npreciopromedio & ")"
      CnAlmacen.Execute Isql
   Else
      RsSubStock.Requery
      RsSubStock.Filter = "tCodigoProducto='" & Producto & "' and tCodigoSubArea='" & Area & "'"
       
      If RsSubStock.RecordCount = 0 Then
         nStock = 0
         CnAlmacen.Execute "Insert into TSUBSTOCK ( tCodigoSubArea, tCodigoProducto, nCorrelativo, nStockActual) " & _
                           "Values ('" & Area & "', '" & Producto & "', 1," & nStock - (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ")"
      Else
         nStock = RsSubStock!nStockActual
         CnAlmacen.Execute "Update TSUBSTOCK set nStockActual = nStockActual - " & (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", ncorrelativo = ncorrelativo + 1 where tCodigoSubArea='" & Area & "' and tCodigoProducto ='" & Producto & "'"
      End If
      
      nCorrela = Calcular("select max(nCorrelativo) as Codigo from MSUBKARDEX where tCodigosubArea='" & Area & "' and tCodigoProducto='" & Producto & "'", CnAlmacen)
      Isql = "Insert into MSUBKARDEX " & _
             "(tCodigoSubArea, tCodigoProducto, nCorrelativo, fRegistro, tUsuario, tTipoDocumento, " & _
             "tDocumento, nIngreso, nSalida, nValor, nStockActual, nStockUltimo, tEstadoKardex) " & _
             "Values ( '" & Area & "', " & _
                      "'" & Producto & "', " & _
                            nCorrela + 1 & ", " & _
                      " getdate(), " & _
                      " '" & sUsuario & "', " & _
                      "'95', " & _
                      "'" & Documento & "', 0, " & _
                            (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            RsProducto!npreciopromedio * (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            nStock - (Cantidad / IIf(TipoDescargo = "D", 1, RsProducto!nFactor)) & ", " & _
                            nStock & ", " & _
                      "'01')"
      CnAlmacen.Execute Isql
   End If
   RsProducto.Requery
   
End Sub

Sub DescargoReceta(Receta As String, Cantidad As Double)
    'Obtengo la Cabecera de la Receta Venta
    RsRecetaVenta.Filter = "tRecetaVenta='" & Receta & "'"
    If RsRecetaVenta.RecordCount > 0 Then
       'Obtengo el Detalle de la Receta Venta
       RsRecetaVentaDetalle.Filter = "tRecetaVenta='" & RsRecetaVenta!tRecetaventa & "'"
       If RsRecetaVentaDetalle.RecordCount > 0 Then
          Do While Not RsRecetaVentaDetalle.EOF
             If RsRecetaVentaDetalle!lInsumo = True Then
                'Descargo los Items Receta Base por insumos
                DescargoInsumo RsRecetaVentaDetalle!tCodigoProducto, RsTransferencia!nCantidad
             Else
                'Descargo los Items Receta Base por Producto y los items normal
                Descargo RsRecetaVenta!tCodigoArea, RsTransferencia!tCodigoPedido, RsRecetaVentaDetalle!tCodigoProducto, RsRecetaVentaDetalle!nCantidad * Cantidad, "R"
             End If
             RsRecetaVentaDetalle.MoveNext
          Loop
       End If
    End If
End Sub

Sub DescargoInsumo(Producto As String, Cantidad As Double)
    Dim RsRecetaBaseDetalle As Recordset
    Dim sRecetaBase As String
    
    'Obtengo la RecetaBase
    sRecetaBase = Calcular("select tRecetaBase as Codigo from mRecetaBase where tCodigoProducto='" & Producto & "'", CnAlmacen)
    If sRecetaBase <> "0" Or sRecetaBase <> 0 Then
       Set RsRecetaBaseDetalle = Lib.OpenRecordset("select * from vRecetaBaseDetalle where tRecetaBase='" & sRecetaBase & "'", CnAlmacen)
       If RsRecetaBaseDetalle.RecordCount > 0 Then
          RsRecetaBaseDetalle.MoveFirst
                    
          Do While Not RsRecetaBaseDetalle.EOF
             If RsRecetaBaseDetalle!lInsumo = True Then
                DescargoInsumo RsRecetaBaseDetalle!tCodigoProducto, Cantidad
             Else
                Descargo RsRecetaBaseDetalle!tCodigoArea, RsTransferencia!tCodigoPedido, RsRecetaBaseDetalle!tCodigoProducto, RsRecetaBaseDetalle!nCantidad * Cantidad, "R"
             End If
             RsRecetaBaseDetalle.MoveNext
          Loop
       End If
       Set RsRecetaBaseDetalle = Nothing
    End If
End Sub

Sub DescargoCombo()
    Isql = "SELECT dbo.CPEDIDO.tCodigoPedido, dbo.CPEDIDO.tProducto, dbo.CPEDIDO.tCodigoProducto, dbo.CPEDIDO.nCantidad AS nCantidad, " & _
           "dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.TPRODUCTO.tArea " & _
           "FROM dbo.CPEDIDO INNER JOIN dbo.TPRODUCTO ON dbo.CPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem AND dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tProducto = dbo.DPEDIDO.tCodigoProducto " & _
           "where dbo.CPEDIDO.tCodigoPedido = '" & RsTransferencia!tCodigoPedido & "' and dbo.CPEDIDO.tItem = '" & RsTransferencia!tItem & "'"
                      
    Set RsComboDetalle = Lib.OpenRecordset(Isql, Cn)
        
    If RsComboDetalle.RecordCount > 0 Then
       RsComboDetalle.MoveFirst
       Do While Not RsComboDetalle.EOF
          Select Case RsComboDetalle!tDescargo
                 Case Is = "D" 'Descargo Directo
                      Descargo RsComboDetalle!tArea, RsTransferencia!tCodigoPedido, RsComboDetalle!tEnlace, RsComboDetalle!nCantidad, "D"
                 Case Is = "R" 'Receta
                      DescargoReceta RsComboDetalle!tEnlace, RsComboDetalle!nCantidad
          End Select
          RsComboDetalle.MoveNext
       Loop
    End If
    Set RsComboDetalle = Nothing
End Sub
