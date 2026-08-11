VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiConsultaIntegrada 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Integral para Restaurantes"
   ClientHeight    =   7065
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   11880
   Icon            =   "mdiConsultaIntegrada.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   90
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":0FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":12EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":1B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":3326
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":3642
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":395E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":5EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsultaIntegrada.frx":67CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6765
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8502
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "19/04/2012"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "08:03 a.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   720
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11880
      TabIndex        =   1
      Top             =   0
      Width           =   11880
   End
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Correlativos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCorrelativoPedido 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnuCorrelativoDocumento 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu mnuNotaCredito 
         Caption         =   "Notas de Crédito"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reservas"
      End
      Begin VB.Menu Linea3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuRecibo 
         Caption         =   "Recibos de Egresos"
      End
      Begin VB.Menu mnuIngreso 
         Caption         =   "Recibos de Ingreso"
      End
      Begin VB.Menu Linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTurno 
         Caption         =   "Turnos"
      End
   End
   Begin VB.Menu mnuReporte 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuControl 
         Caption         =   "De Control"
         Visible         =   0   'False
         Begin VB.Menu mnuLiquidacion 
            Caption         =   "Liquidación de Cajero"
         End
         Begin VB.Menu Linea8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPaloteo 
            Caption         =   "Paloteo de Producción"
         End
         Begin VB.Menu mnuPropiedades 
            Caption         =   "Paloteo de Propiedades"
         End
         Begin VB.Menu mnuPaloteoInsumo 
            Caption         =   "Paloteo de Insumos"
         End
         Begin VB.Menu mnuRepEquivalencias 
            Caption         =   "Paloteo de Equivalencias"
         End
         Begin VB.Menu mnuPaloteoOfertas 
            Caption         =   "Paloteo de Ofertas"
         End
         Begin VB.Menu Linea10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComanda 
            Caption         =   "Comandas"
         End
         Begin VB.Menu mnuCorrelativo 
            Caption         =   "Pedidos"
         End
         Begin VB.Menu mnuPropina 
            Caption         =   "Propinas"
         End
         Begin VB.Menu mnuCortesia 
            Caption         =   "Cortesias"
         End
         Begin VB.Menu mnuDescuento 
            Caption         =   "Descuentos"
         End
         Begin VB.Menu mnuRepCtaCte 
            Caption         =   "Cuentas Corrientes"
         End
         Begin VB.Menu mnuClienteDeuda 
            Caption         =   "Cuentas por Cobrar"
         End
         Begin VB.Menu mnuContacto 
            Caption         =   "Contactos"
         End
         Begin VB.Menu Linea11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAnulado 
            Caption         =   "Control de Transacciones"
         End
         Begin VB.Menu mnuDiferencias 
            Caption         =   "Diferencias entre Paloteo y Liquidacion"
         End
      End
      Begin VB.Menu mnuContables 
         Caption         =   "Contables"
         Begin VB.Menu mnuRegistroVenta 
            Caption         =   "Registro de Ventas"
         End
         Begin VB.Menu mnuRegistroCompras 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu mnuPrincipal 
            Caption         =   "Principales Clientes"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCobranza 
            Caption         =   "Cobranzas (fecha documento)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuEstadistico 
         Caption         =   "Estadísticos"
         Begin VB.Menu mnuIntegradosInsumos 
            Caption         =   "Paloteo de Insumos"
         End
         Begin VB.Menu mnuRanking 
            Caption         =   "Ranking de Producción"
         End
         Begin VB.Menu mnuIntegradosMotorizado 
            Caption         =   "Venta Por Motorizados"
         End
         Begin VB.Menu mnuIntegradosMozos 
            Caption         =   "Venta Por Mozos"
         End
         Begin VB.Menu mnuIntegradoTipOPRod 
            Caption         =   "Venta Por Tipo de Producto"
         End
         Begin VB.Menu mnuTiempoDeliveryInte 
            Caption         =   "Analisis de Tiempos Delivery"
         End
         Begin VB.Menu linea5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMovimientosArticulos 
            Caption         =   "Movimientos de Articulos"
         End
         Begin VB.Menu mnuStockValorizado 
            Caption         =   "Stock Inventario Valorizado"
         End
         Begin VB.Menu mnuMozo 
            Caption         =   "Analítico de Productos por Mozos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMotorizado 
            Caption         =   "Analitico de Productos por Motorizados"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFrecuente 
            Caption         =   "Analítico de Clientes Frecuentes"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstPropina 
            Caption         =   "Producción por Mozos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSolicitadosVsEntregados 
            Caption         =   "Insumos Solicitados Vs. Entregados"
         End
         Begin VB.Menu mnuComparativodeConsumoIns 
            Caption         =   "Comparativo de Consumo de Insumos"
         End
         Begin VB.Menu mnuIntegradoConsolidado 
            Caption         =   "Inventario Consolidado"
         End
         Begin VB.Menu linea7 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTiempoSalon 
            Caption         =   "Tiempos en Salon"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTiempoDelivery 
            Caption         =   "Tiempos Delivery"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudiferenciaDelivery 
            Caption         =   "Diferencias de Tiempos Delivery"
            Visible         =   0   'False
         End
         Begin VB.Menu liena1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRotacion 
            Caption         =   "Rotación de Mesas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOcupabilidad 
            Caption         =   "Ocupabilidad de Mesas"
            Visible         =   0   'False
         End
         Begin VB.Menu Linea9 
            Caption         =   "-"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPaloteoComparativo 
            Caption         =   "Paloteo Comparativo"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuResultadoOperativo 
            Caption         =   "Resultados Operativos de Ventas"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuAnalitico 
         Caption         =   "Gerencial"
         Begin VB.Menu mnuFormaPago 
            Caption         =   "Consolidado de Liquidación de Cajero"
         End
         Begin VB.Menu mnuIntegradoCuentasCorrientes 
            Caption         =   "Cuentas Corrientes"
         End
         Begin VB.Menu mnuPaloteoVentaIntegrado 
            Caption         =   "Paloteo de Venta"
         End
         Begin VB.Menu mnuVentaDiaria 
            Caption         =   "Venta Mensual por Canales"
         End
         Begin VB.Menu mnuVentasFechas 
            Caption         =   "Venta Mensual por Fechas"
         End
         Begin VB.Menu mnuVentaComparada 
            Caption         =   "Venta Comparativa Anual"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVentaComparadaMensual 
            Caption         =   "Venta Comparativa Mensual"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVentaTurno 
            Caption         =   "Venta por Turnos"
            Visible         =   0   'False
         End
         Begin VB.Menu L9 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCobranzaFecha 
            Caption         =   "Cobranza Mensual por Fechas"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuTicketera 
      Caption         =   "&Ticketera"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuLiquidacionTicket 
         Caption         =   "Liquidación de Cajero"
      End
      Begin VB.Menu mnuPaloteoTicket 
         Caption         =   "Paloteo"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiConsultaIntegrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(2).Text = "Usuario : " & sUsuario
    If Not lAlmacen Then
       mnuPaloteoInsumo.Enabled = False
       mnuResultadoOperativo.Enabled = False
    End If
           
    If Not lInfhotel Then
       mnuContacto.Visible = False
    End If
    Call Accesos(Me, "04", sUsuario)
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
End Sub

Public Sub Salir()
   sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
   If sino = vbYes Then
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
        '-----------------------------------------------------------
    End If
    End
   End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
        '-----------------------------------------------------------
    End If
End Sub

Private Sub mnuAcerca_Click()
   frmAbout.Show vbModal
End Sub

Private Sub mnuAnulado_Click()
   frmRepAnulado.Show vbModal
End Sub

Private Sub mnuCancelado_Click()
 '  frmCancelado.Show vbModal
End Sub

Private Sub mnuCliente_Click()
  ' frmRepClienteDelivery.Show vbModal
End Sub

Private Sub mnuClienteDeuda_Click()
   frmRepClienteDeuda.Show vbModal
End Sub

Private Sub mnuCobranzaFecha_Click()
   frmRepCobranzaFecha.Show vbModal
End Sub

Private Sub mnuComparativodeConsumoIns_Click()
frmComparativoConsumoInsumos.Show vbModal
End Sub

Private Sub mnuContacto_Click()
   frmRepContacto.Show vbModal
End Sub

Private Sub mnuCorrelativo_Click()
    frmRepPedido.Show vbModal
End Sub

Private Sub mnuCombinaMozo_Click()
  ' FrmRepMozosCombo.Show vbModal
End Sub

Private Sub mnuCobranza_Click()
  ' frmRepCancelacion.Show vbModal
End Sub

Private Sub mnuComanda_Click()
   frmRepComanda.Show vbModal
End Sub

Private Sub mnuCorrelativoDocumento_Click()
  ' frmDocumentoCorrelativo.Show
End Sub

Private Sub mnuCorrelativoPedido_Click()
  ' frmPedidoCorrelativo.Show
End Sub

Private Sub mnucortesia_Click()
    frmRepCortesia.Show vbModal
End Sub

Private Sub mnuCortesiaMozo_Click()
 ' FrmRepMozosCortesia.Show vbModal
End Sub

Private Sub mnuCtaCte_Click()
  '  frmCtaCte.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
 '  frmCuentaCobrar.Show
End Sub

Private Sub mnuDescuento_Click()
   frmRepDescuento.Show vbModal
End Sub

Private Sub mnudiferenciaDelivery_Click()
   frmRepDiferenciaTiempoDelivery.Show vbModal
End Sub

Private Sub mnuDiferencias_Click()
   frmRepDiferencia.Show vbModal
End Sub

Private Sub mnuEstPropina_Click()
   frmRepProduccionMozo.Show vbModal
End Sub

Private Sub mnuFormaPago_Click()
frmRepFormaPagoIntegrado.Show vbModal
End Sub

Private Sub mnuFrecuente_Click()
   frmRepClieFrecuentes.Show vbModal
End Sub

Private Sub mnuIngreso_Click()
   frmReciboIngreso.Show
End Sub

Private Sub mnuIntegradoConsolidado_Click()
    frmInventarioConsolidado.Show vbModal
End Sub

Private Sub mnuIntegradoCuentasCorrientes_Click()
frmRepCtaCteIntegrado.Show vbModal
End Sub

Private Sub mnuIntegradosInsumos_Click()
    frmRepInsumoVentasIntegrado.Show vbModal
End Sub

Private Sub mnuIntegradosMotorizado_Click()
  frmRepAnaliticoMotorizadoIntegrado.Show vbModal
End Sub

Private Sub mnuIntegradosMozos_Click()
  frmRepAnaliticoMozoIntegrado.Show vbModal

End Sub

Private Sub mnuIntegradoTipOPRod_Click()
    frmRepVentaTipoProductoIntegrado.Show vbModal
End Sub

Private Sub mnuLiquidacion_Click()
   frmRepLiquidacion.Show vbModal
End Sub

Private Sub mnuLiquidacionTicket_Click()
   frmRepLiquidacionTicket.Show vbModal
End Sub

Private Sub mnuMotorizado_Click()
  frmRepAnaliticoMotorizado.Show vbModal
End Sub

Private Sub mnuMovimientosArticulos_Click()
    FrmRepMovimientosIntegrado.Show vbModal
End Sub

Private Sub mnuMozo_Click()
   frmRepAnaliticoMozo.Show vbModal
End Sub

Private Sub mnuNotaCredito_Click()
  frmNotaCredito.Show
End Sub

Private Sub mnuOcupabilidad_Click()
   frmRepOcupabilidad.Show vbModal
End Sub

Private Sub mnuPaloteo_Click()
  frmRepPaloteo.Show vbModal
End Sub

Private Sub mnuPaloteoCombo_Click()
   'frmRepPaloteoCombo.Show vbModal
End Sub

Private Sub mnuPaloteoCortesia_Click()
 '  frmRepPaloteoCortesia.Show vbModal
End Sub

Private Sub mnuPaloteoVenta_Click()
  ' frmRepPaloteoVenta.Show vbModal
End Sub

Private Sub mnuProduccionMozo_Click()
 ' frmRepMozoProduccion.Show vbModal
End Sub

Private Sub mnuPedido_Click()
   frmRepPedido.Show vbModal
End Sub

Private Sub mnuPaloteoComparativo_Click()
   frmRepPaloteoComparativo.Show vbModal
End Sub

Private Sub mnuPaloteoInsumo_Click()
   frmRepInsumoVentas.Show vbModal
End Sub

Private Sub mnuPaloteoOfertas_Click()
   frmRepPaloteoOfertas.Show vbModal
End Sub

Private Sub mnuPaloteoTicket_Click()
   frmRepPaloteoTicket.Show vbModal
End Sub

Private Sub mnuPaloteoVentaIntegrado_Click()
    frmRepPaloteoVentaIntegrado.Show vbModal
End Sub

Private Sub mnuPrincipal_Click()
   frmRepPrincipal.Show vbModal
End Sub

Private Sub mnuPropiedades_Click()
  frmRepPaloteoPropiedades.Show vbModal
End Sub

Private Sub mnuPropina_Click()
   frmRepPropina.Show vbModal
End Sub

Private Sub mnuRankingCombinacion_Click()
  ' frmRepRankingCombo.Show vbModal
End Sub

Private Sub mnuRankingCortesia_Click()
   'frmRepRankingCortesia.Show vbModal
End Sub

Private Sub mnuRankingProduccion_Click()
  ' frmRepRankingProduccion.Show vbModal
End Sub

Private Sub mnuRankingVenta_Click()
  ' frmRepRankingVenta.Show vbModal
End Sub

Private Sub mnuRanking_Click()
  frmRepRankingIntegrado.Show vbModal
End Sub

Private Sub mnuRecibo_Click()
   frmReciboEgreso.Show
End Sub

Private Sub mnuRegistroCompras_Click()
    frmRepRegistroComprasIntegrado.Show vbModal
End Sub

Private Sub mnuRegistroVenta_Click()
   frmRepRegistroVentaIntegrado.Show vbModal
End Sub

Private Sub mnuRepCtaCte_Click()
   frmRepCtaCte.Show vbModal
End Sub

Private Sub mnuRepEquivalencias_Click()
   frmRepPaloteoSubProd.Show vbModal
End Sub

Private Sub mnuReserva_Click()
  ' frmReserva.Show
End Sub

Private Sub mnuResultadoOperativo_Click()
   frmRepResultadoOperativo.Show vbModal
End Sub

Private Sub mnuRotacion_Click()
   frmRepRotacionMesa.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub mnuSolicitadosVsEntregados_Click()
    frmRepSolicitadoVsEntregado.Show vbModal
End Sub

Private Sub mnuStockValorizado_Click()
    FrmRepInventariosIntegrado.Show vbModal
End Sub

Private Sub mnuTiempoDelivery_Click()
   frmRepTiempoDelivery.Show vbModal
End Sub

Private Sub mnuTiempoDeliveryInte_Click()
    frmRepTiempoDeliveryIntegrado.Show vbModal
    
End Sub

Private Sub mnuTiempoSalon_Click()
   frmRepTiempoSalon.Show vbModal
End Sub

Private Sub mnuTurno_Click()
   frmLiquidacion.Show
End Sub

Private Sub mnuVenta_Click()
  ' frmRepVentaAcumulada.Show vbModal
End Sub

Private Sub mnuVentaMozo_Click()
 ' FrmRepMozosVentas.Show vbModal
End Sub

Private Sub mnuVentaComparada_Click()
  frmRepVentaCompAnual.Show vbModal
End Sub

Private Sub mnuVentaComparadaMensual_Click()
   frmRepVentaCompMensual.Show vbModal
End Sub

Private Sub mnuVentaDiaria_Click()
    frmRepVentaMensualCanalesIntegrado.Show vbModal
End Sub

Private Sub mnuVentasFechas_Click()
    frmRepDiarioVentaIntegrado.Show vbModal
End Sub

Private Sub mnuVentaTurno_Click()
 '  frmRepVentasTurno.Show vbModal
End Sub

