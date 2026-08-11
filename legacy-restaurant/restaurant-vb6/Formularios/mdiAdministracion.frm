VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdiAdministracion 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Integral para Restaurantes"
   ClientHeight    =   9255
   ClientLeft      =   1830
   ClientTop       =   -375
   ClientWidth     =   10680
   Icon            =   "mdiAdministracion.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStretch 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   64800
      Left            =   0
      Picture         =   "mdiAdministracion.frx":57E2
      ScaleHeight     =   64800
      ScaleWidth      =   10680
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   10680
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   10200
         Top             =   7920
      End
      Begin VB.Timer Timerwebapp 
         Interval        =   10000
         Left            =   13920
         Top             =   840
      End
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   720
      Top             =   6960
   End
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   10620
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      Begin VB.CommandButton cmdOpcion9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Migrar Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7830
         Picture         =   "mdiAdministracion.frx":11E403
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   6520
         Picture         =   "mdiAdministracion.frx":11F245
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion6 
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   10800
         Picture         =   "mdiAdministracion.frx":11F38F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion7 
         BackColor       =   &H00FFFFFF&
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
         Height          =   660
         Left            =   9140
         Picture         =   "mdiAdministracion.frx":11F4D9
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   5220
         Picture         =   "mdiAdministracion.frx":11F5CB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3920
         Picture         =   "mdiAdministracion.frx":11F715
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mesas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2610
         Picture         =   "mdiAdministracion.frx":11F80F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1300
         Picture         =   "mdiAdministracion.frx":11F909
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Platos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   0
         Picture         =   "mdiAdministracion.frx":11FA0B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1290
      End
      Begin VB.Image ImagePais 
         Height          =   585
         Left            =   12600
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   135
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   8955
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   3793
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "13/01/2026"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "05:21 p.m."
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
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "&Configuracion"
      Begin VB.Menu mnuParametro 
         Caption         =   "Parámetros &Generales"
      End
      Begin VB.Menu linea13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoPedido 
         Caption         =   "Canales de Venta"
      End
      Begin VB.Menu line98 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImpresora 
         Caption         =   "Mantenimiento de &Impresoras"
      End
      Begin VB.Menu mnuConfiguraCaja 
         Caption         =   "Configuración de &Cajas"
      End
      Begin VB.Menu mnuConfiguracionImpresora 
         Caption         =   "Configuración de Impresoras"
      End
      Begin VB.Menu mnuMantSectorVentas 
         Caption         =   "Mantenimiento de Agrupacion de Puntos de Ventas"
      End
      Begin VB.Menu mnuManteEstable 
         Caption         =   "Mantenimiento de Establecimientos (Locales)"
      End
      Begin VB.Menu mnuMensaje 
         Caption         =   "Configura Mensaje"
      End
      Begin VB.Menu mnuSucursales 
         Caption         =   "Sucursales"
      End
      Begin VB.Menu mnuCodigoControl 
         Caption         =   "Configuración Código Control"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCierre 
         Caption         =   "Cierre de &Periodo"
      End
      Begin VB.Menu line99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTablaReplica 
         Caption         =   "Actualización de Tablas"
      End
      Begin VB.Menu xLinea11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuario 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuGrupoUsuario 
         Caption         =   "&Grupo de Usuarios"
      End
      Begin VB.Menu mnuTipoCambio 
         Caption         =   "&Tipos de Cambio"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuTabla 
      Caption         =   "&Tablas"
      Begin VB.Menu mnuIdentidad 
         Caption         =   "Tipos de Identidad"
      End
      Begin VB.Menu mnuTipoDocumento 
         Caption         =   "Tipos de &Documento"
      End
      Begin VB.Menu mnuCancelacion 
         Caption         =   "Otros Tipos de cancelación"
      End
      Begin VB.Menu mnuCuentaContable 
         Caption         =   "Cuentas Contables de Cancelacion"
      End
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoClienteFacturado 
         Caption         =   "Tipos de Clientes Facturados"
      End
      Begin VB.Menu mnuTipoCtaCte 
         Caption         =   "Tipos de Cuenta Corriente"
      End
      Begin VB.Menu mnuTipoCliente 
         Caption         =   "Tipos de Clientes Frecuentes"
      End
      Begin VB.Menu mnuEstadoClienteFrecuente 
         Caption         =   "Estado de Clientes Frecuentes"
      End
      Begin VB.Menu mnuDelivery 
         Caption         =   "Clientes &Frecuentes"
      End
      Begin VB.Menu mnuClienteFactura 
         Caption         =   "Clientes Facturados"
      End
      Begin VB.Menu mnuTransportista 
         Caption         =   "Datos de la Unidad de Transporte y Conductor"
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaitre 
         Caption         =   "Maitres"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMozo 
         Caption         =   "Me&sero"
      End
      Begin VB.Menu mnuMotorizados 
         Caption         =   "Mo&torizados"
      End
      Begin VB.Menu mnuEmpacador 
         Caption         =   "Empacadores"
      End
      Begin VB.Menu mnuZona 
         Caption         =   "&Zonas"
      End
      Begin VB.Menu mnuDistritos 
         Caption         =   "Distritos"
      End
      Begin VB.Menu mnuDispensadores 
         Caption         =   "Dispensadores"
      End
      Begin VB.Menu mnuTarjetaAproximidad 
         Caption         =   "Tarjeta de Proximidad"
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesas 
         Caption         =   "&Mesas"
         Shortcut        =   ^M
      End
      Begin VB.Menu linea5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCortesia 
         Caption         =   "Motivos de Cortesías"
      End
      Begin VB.Menu mnuEliminacion 
         Caption         =   "Motivos de Eliminación"
      End
      Begin VB.Menu mnu_MotivoAnula 
         Caption         =   "Motivos de Anulación"
      End
      Begin VB.Menu mnuDescuento 
         Caption         =   "Motivos de Descuentos"
      End
      Begin VB.Menu mnuMotivoReserva 
         Caption         =   "Motivos de Reserva"
      End
      Begin VB.Menu mnuTipoEgreso 
         Caption         =   "Tipos de Egreso"
      End
      Begin VB.Menu linea12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTarjetaCredito 
         Caption         =   "Tarjetas &Bancarias"
      End
      Begin VB.Menu mnuArea 
         Caption         =   "&Areas de Producción"
      End
      Begin VB.Menu mnuCCosto 
         Caption         =   "Centro de Costo"
      End
      Begin VB.Menu mnuTerminales 
         Caption         =   "Terminales"
      End
   End
   Begin VB.Menu mnuProd 
      Caption         =   "&Productos de Venta"
      Begin VB.Menu mnuTipoProducto 
         Caption         =   "&Tipos de Producto"
      End
      Begin VB.Menu mnuUnidadNEgocio 
         Caption         =   "Unidad de Negocios"
      End
      Begin VB.Menu linea11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperador 
         Caption         =   "Operadores"
      End
      Begin VB.Menu mnuPropiedad 
         Caption         =   "P&ropiedades"
         Shortcut        =   ^R
      End
      Begin VB.Menu linea9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "&Grupos y SubGrupos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuProducto 
         Caption         =   "&Productos y Precios"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuAgrupacion 
         Caption         =   "Agrupaciones de Caja Rápida"
      End
      Begin VB.Menu mnuProgramacionPrecios 
         Caption         =   "Programación de Cambio de  Precios"
      End
      Begin VB.Menu mnuInsumoCritico 
         Caption         =   "&Insumos/Platos de Stock Crítico"
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOferta 
         Caption         =   "&Mantenimiento de Ofertas y Recomendaciones"
      End
   End
   Begin VB.Menu mnuUtilitario 
      Caption         =   "&Utilitarios"
      Begin VB.Menu mnuOptimizador 
         Caption         =   "&Optimizador de la BD"
      End
      Begin VB.Menu linea10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Visible         =   0   'False
      End
      Begin VB.Menu linea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlmacen 
         Caption         =   "Descargo de Ventas"
      End
      Begin VB.Menu mnuCambiaLocal 
         Caption         =   "Cambiar de Local"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuContabilidad 
         Caption         =   "Transferencia Contabilidad"
         Visible         =   0   'False
      End
      Begin VB.Menu Linea4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRTemporal 
         Caption         =   "Restore Temporal"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTemporal 
         Caption         =   "Actualiza Temporal"
         Visible         =   0   'False
      End
      Begin VB.Menu mnulinea20 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuElimina 
         Caption         =   "Eliminar Cortesias"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuActualizar 
         Caption         =   "Actualizar Sistema"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu mnuLiberar 
         Caption         =   "Liberar Descargo Venta"
      End
      Begin VB.Menu mnuLiberarCaja 
         Caption         =   "Liberar Caja Punto de Venta"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiAdministracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'PRUEBA TIMER
' Variables a nivel de formulario (persisten entre ticks)
Private secondsLeft As Long
Private timerRunning As Boolean

'PRUEBA TIMER

Private Const IMAGESIZE = 0.566893424036281

Private Sub cmdOpcion1_Click()
   frmProducto.Show
End Sub

Private Sub cmdOpcion2_Click()
   frmClienteDelivery.Show
End Sub

Private Sub cmdOpcion3_Click()
   frmMantenimientoMesa.Show
End Sub

Private Sub cmdOpcion4_Click()
   frmUsuario.Show
End Sub

Private Sub cmdOpcion5_Click()
   frmBackup.Show
End Sub

Private Sub cmdOpcion7_Click()
   salir
End Sub

Private Sub cmdOpcion6_Click()
    sTipo = "Local"
    frmRestore.Show
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    picStretch.Visible = False
    picStretch.AutoRedraw = True
    picStretch.Height = Me.Height
    picStretch.Width = Me.Width
    ImageWidth = picStretch.Picture.Width * IMAGESIZE
    ImageHeight = (picStretch.Picture.Height * IMAGESIZE) + 3000
    picStretch.PaintPicture picStretch.Picture, 0, 0, Me.Width, Me.Height, 0, 0, ImageWidth, ImageHeight
    Set Me.Picture = picStretch.Image
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
        'auditoriA
         registroAccesoAuditoria "S", sUsuario
        'auditoria

End Sub

Private Sub mnuCodAutorizacion_Click()
    frmCodigoAutorizacion.Show
End Sub

Private Sub mnu_MotivoAnula_Click()
   frmMotivoAnulacion.Show
End Sub

Private Sub mnuActualizar_Click()
    Dim val As String
    'Call inputbox_password(Me, "*")
    val = InputBox("Ingrese clave de acceso", "Seguridad")
    
    If (Trim(val) = "4gust1n-fl0r14n") Then
        Dim verificaForms As Boolean
        verificaForms = VerificaFormAbiertos
        If verificaForms <> True Then
            frmActualizar.Show
        Else
            MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
        End If
    End If
End Sub

Private Sub mnuCCosto_Click()
    frmCentroCosto.Show
End Sub

Private Sub mnuCodigoControl_Click()
    frmCajaCodigoControl.Show
End Sub
Private Sub cmdOpcion8_Click()
    Dim verificaForms As Boolean
        verificaForms = VerificaFormAbiertos
        If verificaForms <> True Then
            FrmUpdate.Show
        Else
            MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
        End If
End Sub


Private Sub cmdOpcion9_Click()
  
    
    Dim verificaForms As Boolean
    verificaForms = VerificaFormAbiertos
    If verificaForms <> True Then
        FrmUpdateAlmacen.Show
    Else
        MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
    End If
End Sub

Private Sub MDIForm_Load()

    Timer1.Enabled = False
    timerRunning = False
    InitializeTimerFromParametro

    StatusBar.Panels.Item(1).Text = "Local: " & localConectado
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & sUsuario
    cmdOpcion5.Visible = lBotonTrans
    
    Centrar Me
    If lCentral = False Then
        line99.Visible = False
        mnuTablaReplica.Visible = False
        cmdOpcion8.Visible = False
    End If
 
    If lAlmacen Then
       mnuAlmacen.Visible = True
       Linea4.Visible = True
    Else
       mnuAlmacen.Visible = False
       Linea4.Visible = False
    End If
       
    If lAlmacenRemoto = True Then
         cmdOpcion9.Visible = True
    Else
         cmdOpcion9.Visible = False
    End If
    
    If lInfhotel Then
       linea11.Visible = False
       'mnuUsuario.Visible = False
       'cmdOpcion4.Visible = False
       mnuTipoCambio.Visible = False
    End If
    
    
    If lSAP Then
        mnuTipoClienteFacturado.Visible = True
    Else
        mnuTipoClienteFacturado.Visible = False
    End If
    
     Call Accesos(mdiAdministracion, "03", sUsuario)
        If lCentral = True Then
              Call accesoCentralizada(Me, "03")
              cmdOpcion2.Enabled = False
            mnuTipoCtaCte.Enabled = False
        End If
    If lMCPV Then
        mnuMozo.Visible = False
    End If
    If multiLocal = False Then
        mnuCambiaLocal.Visible = False
    End If
    
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
    
        Select Case pais
           Case "001" 'Bolivia
                mnuCodigoControl.Visible = True
                 mnuManteEstable.Visible = True
           Case "002" 'Ecuador
                mnuSucursales.Visible = False
                mnuManteEstable.Visible = True
           Case "005" 'España  JAPN 10-04-2025
                mnuTipoCambio.Visible = False
           Case Else
                mnuCodigoControl.Visible = False
                mnuManteEstable.Visible = True
    End Select
    
    Select Case sEmpresa
           Case Is = "002"
                mnuContabilidad.Visible = True
           Case Is = "004"
                mnuRTemporal.Visible = True
                mnuTemporal.Visible = True
                mnulinea20.Visible = True
    End Select
    
'    If pais = "005" Then
'        mnuMozo.Caption = "Camarero"
'    Else
'        mnuMozo.Caption = "Mesero"
'    End If

    mnuMozo.Caption = Mesero
    
    'IMAGEN PAIS
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(App.Path & "\bmps\Paises\" & pais & ".jpg") Then
               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\" & pais & ".jpg")
    End If
    
    Call IniciarNoticia
    
'    Select Case pais
'        Case Is = "001"
'            If fso.FileExists(App.Path & "\bmps\Paises\001.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\001.jpg")
'            End If
'        Case Is = "002"
'            If fso.FileExists(App.Path & "\bmps\Paises\002.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\002.jpg")
'            End If
'        Case Else
'            If fso.FileExists(App.Path & "\bmps\Paises\000.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\000.jpg")
'            End If
'    End Select
    
'    If fso.FileExists(App.Path & "\bmps\Pais.jpg") Then
'       ImagePais.Picture = LoadPicture(App.Path & "\bmps\Pais.jpg")
'    End If
    Set fso = Nothing
    
End Sub
Public Function obtieneNoticia() As String
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("SP_NOTICIAS_INFOREST", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 1
                  oComando.CreateParameter "@usuario", adVarChar, adParamInput, 50, sUsuario
                  oComando.CreateParameter "@resultado", adVarChar, adParamOutput, 200, ""
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneNoticia = oComando.GetParameterValue("@resultado")
End Function
Public Sub IniciarNoticia()
On Error GoTo fin:
    Dim resul As String
    resul = obtieneNoticia
    If resul = "1" Then
        Exit Sub
    End If


    Dim envio As String
    'envio = sUsuario & "," & "http://www.licencias.infomatica.pe/assets/image/WhatsApp Image 2021-10-21 at 15.24.13.jpeg"
    envio = sUsuario & "," & resul & "," & sMDB & "," & sRuta

    'INICIO DE ACTUALIZADOR
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.Path & "\Noticia\InfoNoticias.exe") Then
        'Kill App.Path & "\images\fact.bmp"
        ShellExecute 0, "Open", App.Path & "\Noticia\InfoNoticias.exe", envio, "", 1
'    Else
'        MsgBox "La ruta de la Noticia no se ha encontrado, favor de informar al administrador de sistemas!!!: "
    End If
    
   ''''  Shell App.path & "\Noticia\InfoNoticias.exe", vbNormalFocus
Exit Sub
fin:
MsgBox "Noticia: " & Error
End Sub
Public Sub salir()
   ' If MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje) = vbYes Then
        If lHARDkey Then
            '--------Libera Licencia de la Llave------------------------
            Dim Verifica As Boolean
            Verifica = hk.FinalizarConexion(Aplicacion.Administracion)
            '-----------------------------------------------------------
            
            
        End If
        
               
        'auditoriA
         registroAccesoAuditoria "S", sUsuario
        'auditoria
      End
  'End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.Administracion)
        '-----------------------------------------------------------
    End If
            'auditoriA
         registroAccesoAuditoria "S", sUsuario
        'auditoria

End Sub

Private Sub mnuAcerca_Click()
   frmAbout.Show vbModal
End Sub

Private Sub mnuAgrupacion_Click()
   frmAgrupacion.Show
End Sub

Private Sub mnuAlmacen_Click()

 Dim RsDescargoAutomatico As New Recordset
   Set RsDescargoAutomatico = Lib.OpenRecordset("select lDescargoAutomatico from tparametro", Cn)
   If IIf(IsNull(RsDescargoAutomatico!lDescargoAutomatico), False, RsDescargoAutomatico!lDescargoAutomatico) Then
     MsgBox "No se puede usar el descargo manual porque el descargo automático está activo.", vbInformation, sMensaje
     Exit Sub
   End If
   
   
   frmDescargo.Show vbModal
End Sub

Private Sub mnuArea_Click()
  frmArea.Show
End Sub

Private Sub mnuBackup_Click()
   frmBackup.Show
End Sub

Private Sub mnuCambiaLocal_Click()
    Dim verificaForms As Boolean
        verificaForms = VerificaFormAbiertos
        If verificaForms <> True Then
            frmServidorEnlace.Show
        Else
            MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
        End If
End Sub

Private Sub mnuCancelacion_Click()
   frmTipoCancelacion.Show
End Sub

Private Sub mnuCierre_Click()
   frmCierrePeriodo.Show
End Sub

Private Sub mnuClienteFactura_Click()
   frmClienteFactura.Show
End Sub

Private Sub mnuConfiguraCaja_Click()
   frmCaja.Show
End Sub

Private Sub mnuConfiguracionImpresora_Click()
    On Error Resume Next
    dlgPrinter.ShowPrinter
End Sub

Private Sub mnuCortesia_Click()
   frmCortesia.Show
End Sub


Private Sub mnuCuentaContable_Click()
   frmCuentaPago.Show vbModal
End Sub

Private Sub mnuDelivery_Click()
   frmClienteDelivery.Show
End Sub

Private Sub mnuDescuento_Click()
   frmMotivoDescuento.Show
End Sub

Private Sub mnuDispensadores_Click()
 FrmDispensadores.Show
End Sub

Private Sub mnuDistritos_Click()
   frmDistrito.Show
End Sub

Private Sub mnuElimina_Click()
   wEnter = False
   
   'Cambia el MsgBox
   If MsgBox("Seguro de Eliminar los movimientos?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
      Exit Sub
   End If
   
   frmPassword.Show vbModal
   If wEnter Then
      If sDescrip <> sPAdmin Then
         MsgBox "Password incorrecto", vbCritical, sMensaje
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   
   'Cambia el MsgBox
   If MsgBox("Si prosigue se eliminará todos los movimientos", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Cn.Execute "delete from DDOCUMENTO where tDocumento in (select tDocumento from MDOCUMENTO where tTipoDocumento='00')"
   Cn.Execute "delete from MDOCUMENTO where tTipoDocumento ='00'"
   
   Screen.MousePointer = vbDefault
   MsgBox "La Eliminación se realizó con exito", vbInformation, sMensaje

End Sub

Private Sub mnuEliminacion_Click()
   frmMotivoEliminacion.Show
End Sub

Private Sub mnuEmpacador_Click()
   frmEmpacadora.Show
End Sub

Private Sub mnuEstadoClienteFrecuente_Click()
    frmEstadoClienteFrecuente.Show
End Sub

Private Sub mnuGrupo_Click()
   frmGrupo.Show
End Sub

Private Sub mnuGrupoUsuario_Click()
   frmGrupoAcceso.Show
End Sub

Private Sub mnuIdentidad_Click()
   frmTipoIdentidad.Show
End Sub

Private Sub mnuImpresora_Click()
   frmImpresora.Show
End Sub
 
Private Sub mnuInsumoCritico_Click()
    frmInsumo.Show
End Sub

Private Sub mnuLiberar_Click()
 Cn.Execute "UPDATE tparametro set tsolouno=0, lNomSoloUno='' "
 
 MsgBox "!Descargo de Venta Liberado", vbInformation, "Descargo de ventas"
 
 

 
 
End Sub

Private Sub mnuLiberarCaja_Click()
 Cn.Execute "UPDATE LOG_SESIONES SET lActivo = 1 where lActivo = 0 and tcaja = '" & sCaja & "'"
 MsgBox "Caja Liberada " & sCaja, vbInformation, "Sesiones"
End Sub

Private Sub mnuMaitre_Click()
    frmMaitre.Show
End Sub

Private Sub mnuManteEstable_Click()
    frmLocal.Show 'vbModal
End Sub

Private Sub mnuMantSectorVentas_Click()
   frmSectorVenta.Show
End Sub

Private Sub mnuMensaje_Click()
   frmConfiguraMensaje.Show
End Sub

Private Sub mnuMesas_Click()
   frmMantenimientoMesa.Show vbModal
End Sub

Private Sub mnuMotivoReserva_Click()
   frmMotivoReserva.Show
End Sub

Private Sub mnuMotorizados_Click()
   frmMotorizado.Show
End Sub

Private Sub mnuMozo_Click()
   frmMozo.Show
End Sub

Private Sub mnuOferta_Click()
   frmOferta.Show
End Sub

Private Sub mnuOperador_Click()
   frmOperador.Show
End Sub

Private Sub mnuOptimizador_Click()
   Call EliminaTemporal
   FrmOptimizar.Show
End Sub

Private Sub mnuParametro_Click()
    bCargaFormulario = False
   frmParametro.Show vbModal
End Sub

Private Sub mnuProducto_Click()
   frmProducto.Show
End Sub

Private Sub mnuProgramacionPrecios_Click()
   frmProgramacionPrecios.Show
End Sub

Private Sub mnuPropiedad_Click()
   frmPropiedad.Show
End Sub

Private Sub mnuRestore_Click()
   sTipo = "Local"
   frmRestore.Show
End Sub

Private Sub mnuRTemporal_Click()
   sTipo = "Temporal"
   frmRestore.Show
End Sub

Private Sub mnuSalir_Click()
   salir
End Sub

Private Sub mnuSucursales_Click()
    frmSucursal.Show
End Sub

Private Sub mnuTablaReplica_Click()
    frmGrupoReplica.Show
End Sub

Private Sub mnuTarjetaAproximidad_Click()
FrmTarjetaAproximidad.Show
End Sub

Private Sub mnuTarjetaCredito_Click()
   frmTarjeta.Show
End Sub

Private Sub mnuTemporal_Click()
   Dim xInicio As Date
   Dim CnTemp As Connection
   Dim RsFecha As Recordset
   Dim fFechaI As Date
   Dim fFechaF As Date
   Screen.MousePointer = vbHourglass
   mdiAdministracion.StatusBar.Panels.Item(2).Text = "Actualizando Movimiento de Pedidos"
   
   Set CnTemp = New Connection
   CnTemp.Provider = "SQLOLEDB"
   CnTemp.CursorLocation = adUseServer
   CnTemp.ConnectionString = "User ID=infhotel" & _
                             ";password=4gust1n-fl0r14n" & _
                             ";Data Source=" & sRuta & _
                             ";Initial Catalog=TEMPORAL"
   CnTemp.CommandTimeout = 300
   CnTemp.Open
   
   Set RsFecha = Lib.OpenRecordset("select top 1 * from menvio where lcierre = 1 order by fFinal desc", CnTemp)
   If RsFecha.EOF Then
      MsgBox "Error: No ha realizado el Cierre", vbCritical, sMensaje
      Exit Sub
   Else
      fFechaI = RsFecha!fInicio
      fFechaF = RsFecha!fFinal
   End If
         
   'MENVIO
   Cn.Execute "delete from MENVIO"
   Cn.Execute "INSERT into MENVIO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.MENVIO"
   
   'Anulacion Pedidos
   Cn.Execute "delete from APEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into APEDIDO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.APEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   'Combos Pedidos
   Cn.Execute "delete CPEDIDO " & _
              "FROM dbo.CPEDIDO LEFT OUTER JOIN dbo.MPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
              "where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into CPEDIDO select T1.* from " & _
              "OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.CPEDIDO T1 LEFT OUTER JOIN " & _
              "OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.MPEDIDO T2 ON " & _
              "T1.tCodigoPedido = T2.tCodigoPedido where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
                 
   'Detalle Pedidos
   Cn.Execute "delete from DPEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into DPEDIDO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.DPEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   'Cabecera Pedidos
   Cn.Execute "delete from MPEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into MPEDIDO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.MPEDIDO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   'Cabecera Documentos
   mdiAdministracion.StatusBar.Panels.Item(2).Text = "Actualizando Movimiento de Documentos"
   Cn.Execute "delete from MDOCUMENTO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into MDOCUMENTO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.MDOCUMENTO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   'Detalle Documentos
   Cn.Execute "delete DDOCUMENTO " & _
              "FROM dbo.DDOCUMENTO LEFT OUTER JOIN dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
              "where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into DDOCUMENTO select T1.* from " & _
              "OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.DDOCUMENTO T1 LEFT OUTER JOIN " & _
              "OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.MDOCUMENTO T2 ON " & _
              "T1.tDocumento = T2.tDocumento where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   'Pago Documentos
   Cn.Execute "delete from DPAGODOCUMENTO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   Cn.Execute "INSERT into DPAGODOCUMENTO select * from OPENDATASOURCE('SQLOLEDB','DataSource=" & sRuta & ";User Id=sa;Password=sistemas').TEMPORAL.dbo.DPAGODOCUMENTO where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & " 23:59'"
   
   StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
   Screen.MousePointer = vbDefault
   MsgBox "Proceso realizado satisfactoriamente", vbInformation, sMensaje
End Sub

Private Sub mnuTerminales_Click()
   frmTerminal.Show
End Sub

Private Sub mnuTipoCambio_Click()
   frmTipoCambio.Show
End Sub

Private Sub mnuTipoCliente_Click()
   frmTipoCliente.Show
End Sub

Private Sub mnuTipoClienteFacturado_Click()
    frmTipoGrupoCliente.Show
End Sub

Private Sub mnuTipoCtaCte_Click()
   frmTipoCtaCte.Show
End Sub

Private Sub mnuTipoDocumento_Click()
   frmTipoDocumento.Show
End Sub

Private Sub mnuTipoEgreso_Click()
frmTipoEgreso.Show
End Sub

Private Sub mnuTipoPedido_Click()
frmTipoPedido.Show
End Sub

Private Sub mnuTipoProducto_Click()
  frmTipoProducto.Show
End Sub

Private Sub mnuTransportista_Click()
   frmDatoTransportista.Show
End Sub

Private Sub mnuUnidadNEgocio_Click()
   frmUnidadNegocio.Show
End Sub

Private Sub mnuUsuario_Click()
   frmUsuario.Show
End Sub

Private Sub mnuZona_Click()
   frmZona.Show
End Sub

Public Sub accesoCentralizada(formulario As Form, tModulo As String)
    Dim Rs As Recordset
    Dim sGrupo As String
    Dim sControl As String
    Dim xObj As Object
    Dim xObjB As Object
    'Isql = "select tNombreObjeto FROM dbo.TACCESO INNER JOIN  dbo.TGRUPOACCESO ON dbo.TACCESO.tCodigoAcceso = dbo.TGRUPOACCESO.tCodigoAcceso " & _
           "where  TACCESO.tModulo='" & tModulo & "' and tFormulario='" & Formulario.Name & "' aND (dbo.TACCESO.tCodigoAcceso   IN   (SELECT tcodigoacceso FROM  taccesoenvia))"
    Isql = "SELECT     dbo.TACCESO.tNombreObjeto, ISNULL(dbo.TACCESOENVIA.tCodigoAcceso, '0') AS envia FROM         dbo.TACCESO INNER JOIN                       dbo.TGRUPOACCESO ON dbo.TACCESO.tCodigoAcceso = dbo.TGRUPOACCESO.tCodigoAcceso LEFT OUTER JOIN                       dbo.TACCESOENVIA ON dbo.TACCESO.tCodigoAcceso = dbo.TACCESOENVIA.tCodigoAcceso where tacceso.tTIPOobjeto='MN' AND  TACCESO.tModulo='" & tModulo & "' and tFormulario='" & formulario.Name & "'"
    ' Debug.Print Isql
    Set Rs = Lib.OpenRecordset(Isql, Cn)
    Do While Not Rs.EOF
        If Rs.Fields("envia") <> "0" Then
           Set xObj = formulario.Controls(Rs!tnombreObjeto)
       
            If Rs!tnombreObjeto = "mnuUsuario" Then
                    Set xObjB = cmdOpcion4
                        xObjB.Enabled = False
            End If
            
            If Rs!tnombreObjeto = "mnuProducto" Then
                    Set xObjB = cmdOpcion1
                        xObjB.Enabled = False
            End If
            
            xObj.Enabled = False
        Else
            
           Set xObj = formulario.Controls(Rs!tnombreObjeto)
           Dim Estado As Boolean
           Estado = xObj.Enabled
            If Estado = False Then
                If Rs!tnombreObjeto = "mnuUsuario" Then
                         Set xObjB = cmdOpcion4
                             xObjB.Enabled = True
                 End If
                   If Rs!tnombreObjeto = "mnuProducto" Then
                         Set xObjB = cmdOpcion1
                             xObjB.Enabled = True
                 End If
           ' MsgBox verificar
          '  xObj.Enabled = True
            End If
        End If
            Rs.MoveNext
        
    Loop
End Sub
'UN EXE VARIAS BD
Public Sub reinicia()
    Unload Me
    mdiAdministracion.Show
End Sub

Private Sub Timer_Timer()
    If multiLocal = True Then
        If ultimoConectado = False Then
             frmServidorEnlace.Show
            Timer.Interval = 0
        End If
    End If
End Sub

Private Sub TimerActualizador_Timer()
On Error GoTo fin:
'    Dim CnSeg As New Connection
'
'    Set CnSeg = New Connection
'    CnSeg.Provider = "SQLOLEDB"
'    CnSeg.CursorLocation = adUseServer
'    CnSeg.ConnectionString = "User ID=" & sUserName & _
'                          ";password=" & sPassword & _
'                          ";Data Source=" & sRuta & _
'                          ";DataTypeCompatibility=80" & _
'                          ";Initial Catalog=INFSEGURIDAD"
'    CnSeg.CommandTimeout = 250
'    CnSeg.Open

    If ActivoActualizador Then
'        If CnSeg.State = 0 Then
'            CnSeg.Open
'        End If
'
'        Dim Valor, ValorMotor, ValorMotorActual As String
'        Valor = Calcular("select top 1 VersionInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
'        'ValorMotor = Calcular("select top 1 VersionFeInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
'        'ValorMotorActual = Calcular("select top 1 VersionActualFEInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
'        Open App.Path & "\version.txt" For Input As #1   ' Abre el archivo para recibir los datos.
'        Line Input #1, Valor
'        Close #1
'
'        Dim sVersionExe As String
'        sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
'
'        If (ValorMotor <> "") Then
'            If (ValorMotor <> ValorMotorActual) Then
'                If (lFEBiz) Then
'                    FrmActualizacion.tipo = "FE"
'                    FrmActualizacion.lblMensaje.Caption = "Nueva Actualización disponible de Facturacion Electronica" & vbNewLine & " FE Versión: " & ValorMotor
'                    FrmActualizacion.Show vbModal
'                    Exit Sub
'                End If
'            End If
'        End If
        
'        If (Valor <> "") Then
'            If (Valor <> sVersionExe) Then
'                FrmActualizacion.tipo = "INFOREST"
'                FrmActualizacion.lblMensaje.Caption = "Nueva Actualización disponible" & vbNewLine & " InfoRest Versión: " & Valor & vbNewLine & "Se recomienda actualizar el sistema, ya que puede presentar problemas en algunos procesos.!!!"
'                FrmActualizacion.Show vbModal
'                Exit Sub
'            End If
'        End If
    End If
Exit Sub
fin:
MsgBox "Error:" & Error
End Sub

Private Sub Timer1_Timer()

If Not timerRunning Then Exit Sub

    secondsLeft = secondsLeft - 1

    If secondsLeft <= 0 Then
        Timer1.Enabled = False
        timerRunning = False
        On Error Resume Next
        Unload Me    ' cierra este formulario
        End          ' termina la aplicación (cierra sesión)
    End If

End Sub

Private Sub Timerwebapp_Timer()
    Dim fechaServer As String
    Dim Isql As String
    Dim IsqlfechaServer As String
    Dim IsqlCierre As String
    Dim statusHora As Boolean
    Dim RsAlertaVenta As Recordset
    Dim RsStatusCierre As Recordset
    Dim RsFechaServer As Recordset
'    obtener fecha formato YYYY/MM/DD del servidor
    IsqlfechaServer = "SELECT CONVERT (date, SYSDATETIME()) as FechaServer"
'    validaciòn de rango horario en el que se ejecutara la alerta
    statusHora = Calcular("SELECT CASE WHEN CONVERT(time, GETDATE()) BETWEEN '15:30:00' AND '15:59:59' THEN 'True' Else 'False'END AS Codigo", Cn)
'    consulta configuraciòn de alerta venta
    Isql = "select lAlertaVenta from TPARAMETRO"
    Set RsAlertaVenta = Lib.OpenRecordset(Isql, Cn)
    Set RsFechaServer = Lib.OpenRecordset(IsqlfechaServer, Cn)
    fechaServer = RsFechaServer!fechaServer
'    validaciòn de cierre de turno
    IsqlCierre = "SELECT CASE WHEN EXISTS (SELECT 1 From MTURNO WHERE CONVERT(date, fInicial) = '" + fechaServer + "' AND lCierre = 0) THEN 1 ELSE 0 END AS Codigo"
    Set RsStatusCierre = Lib.OpenRecordset(IsqlCierre, Cn)
    If (RsAlertaVenta!lAlertaVenta = True) Then
        If (statusHora) Then
            If (RsStatusCierre!codigo = 1) Then
                MsgBox ("Es necesario realizar descarga de Ventas")
            End If
        End If
    End If
End Sub


Private Sub InitializeTimerFromParametro()
    Dim mins As Long
    bCargaFormulario = True
    ' validar que el formulario de parámetros exista / esté cargado
    On Error GoTo CleanExit
    If frmParametro Is Nothing Then Exit Sub

    ' Validar y leer el valor de minutos
    If Not IsNumeric(Trim$(frmParametro.txtTimer.Text)) Then
        mins = 1
    Else
        mins = CLng(val(frmParametro.txtTimer.Text))
        If mins < 1 Then mins = 1
    End If

    ' Si el checkbox está activado, arrancar el timer
    If frmParametro.chkTimer.Value = vbChecked Or frmParametro.chkTimer.Value = 1 Then
        secondsLeft = mins * 60      ' convertir a segundos
        Timer1.Interval = 1000       ' tick cada segundo
        Timer1.Enabled = True
        timerRunning = True
    Else
        Timer1.Enabled = False
        timerRunning = False
    End If

CleanExit:
    On Error Resume Next
    Exit Sub
End Sub
