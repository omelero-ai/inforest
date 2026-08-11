VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9990
   ClientLeft      =   2010
   ClientTop       =   1890
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   Icon            =   "frmCajaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   11955
   Begin VB.Frame fraDetalle 
      Height          =   9210
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton btnNewModal 
         Caption         =   "Ver más"
         Height          =   735
         Left            =   10440
         TabIndex        =   329
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDetallado 
         Height          =   285
         Left            =   3240
         MaxLength       =   25
         TabIndex        =   142
         Text            =   " "
         Top             =   225
         Width           =   5970
      End
      Begin TabDlg.SSTab tabOpcion 
         Height          =   7170
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12647
         _Version        =   393216
         Tabs            =   12
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Activaciones"
         TabPicture(0)   =   "frmCajaDetalle.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame8"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Obligatoriedades"
         TabPicture(1)   =   "frmCajaDetalle.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "chkComanda"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chkObligaPrinter"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chkCancelacion"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chkObligaPrecuenta"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "chkObservacion"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "chkConsumo1"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "chkConsumo2"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "chkConsumo3"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "chkConsumo4"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Impresiones"
         TabPicture(2)   =   "frmCajaDetalle.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(1)=   "Frame11"
         Tab(2).Control(2)=   "Frame9"
         Tab(2).Control(3)=   "Frame15"
         Tab(2).Control(4)=   "Frame20"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Documentos"
         TabPicture(3)   =   "frmCajaDetalle.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdGrilla"
         Tab(3).Control(1)=   "cmdOpcionGrilla(0)"
         Tab(3).Control(2)=   "cmdOpcionGrilla(1)"
         Tab(3).Control(3)=   "cmdOpcionGrilla(2)"
         Tab(3).Control(4)=   "fraGrilla"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Areas de Impresión"
         TabPicture(4)   =   "frmCajaDetalle.frx":04B2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "grdAI"
         Tab(4).Control(1)=   "cmdOpcionGrilla(6)"
         Tab(4).Control(2)=   "cmdOpcionGrilla(5)"
         Tab(4).Control(3)=   "cmdOpcionGrilla(7)"
         Tab(4).Control(4)=   "fraArea"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "Periféricos Adicionales"
         TabPicture(5)   =   "frmCajaDetalle.frx":04CE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame12"
         Tab(5).Control(1)=   "Frame1"
         Tab(5).Control(2)=   "Frame3"
         Tab(5).Control(3)=   "Frame2"
         Tab(5).Control(4)=   "Frame4"
         Tab(5).Control(5)=   "Frame5"
         Tab(5).ControlCount=   6
         TabCaption(6)   =   "Multi Area de Producción"
         TabPicture(6)   =   "frmCajaDetalle.frx":04EA
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "fra1"
         Tab(6).Control(1)=   "chkMulti1"
         Tab(6).Control(2)=   "fra2"
         Tab(6).Control(3)=   "chkMulti2"
         Tab(6).ControlCount=   4
         TabCaption(7)   =   "Areas Chef Control"
         TabPicture(7)   =   "frmCajaDetalle.frx":0506
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "fraAreaChef"
         Tab(7).Control(1)=   "cmdOpcionGrilla(12)"
         Tab(7).Control(2)=   "cmdOpcionGrilla(13)"
         Tab(7).Control(3)=   "cmdOpcionGrilla(14)"
         Tab(7).Control(4)=   "grdAChef"
         Tab(7).ControlCount=   5
         TabCaption(8)   =   "Imágenes en Documentos"
         TabPicture(8)   =   "frmCajaDetalle.frx":0522
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "cmdQuitarFotoPie"
         Tab(8).Control(1)=   "cmdQuitarFotoCabecera"
         Tab(8).Control(2)=   "cmdAgregarFotoPie"
         Tab(8).Control(3)=   "cmdAgregarFoto"
         Tab(8).Control(4)=   "dlgFoto"
         Tab(8).Control(5)=   "dlgFotoPie"
         Tab(8).Control(6)=   "imgFotoPie"
         Tab(8).Control(7)=   "imgFoto"
         Tab(8).ControlCount=   8
         TabCaption(9)   =   "Integraciones"
         TabPicture(9)   =   "frmCajaDetalle.frx":053E
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "FrmFacArgentina"
         Tab(9).Control(1)=   "Frame17"
         Tab(9).Control(2)=   "Frame16"
         Tab(9).ControlCount=   3
         TabCaption(10)  =   "Bloqueos"
         TabPicture(10)  =   "frmCajaDetalle.frx":055A
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "cmdopcion(4)"
         Tab(10).Control(1)=   "ChkBloqueoEdicionPedidoConsolaAgr"
         Tab(10).Control(2)=   "ChkBloqueaPedidoSinMesa"
         Tab(10).Control(3)=   "ChkBloqueaCanal(5)"
         Tab(10).Control(4)=   "ChkBloqueaCanal(4)"
         Tab(10).Control(5)=   "ChkBloqueaCanal(3)"
         Tab(10).Control(6)=   "ChkBloqueaCanal(2)"
         Tab(10).Control(7)=   "ChkBloqueaCanal(1)"
         Tab(10).Control(8)=   "ChkBloqueoAnulaPedido"
         Tab(10).Control(9)=   "ChkBloqueoAnulaItemsPedido"
         Tab(10).Control(10)=   "ChkBloqAfectoInafecto"
         Tab(10).Control(11)=   "ChkBloqEmiDoc"
         Tab(10).Control(12)=   "dgvBloqueoOrigenVenta"
         Tab(10).ControlCount=   13
         TabCaption(11)  =   "+ Activaciones"
         TabPicture(11)  =   "frmCajaDetalle.frx":0576
         Tab(11).ControlEnabled=   0   'False
         Tab(11).Control(0)=   "Frame19"
         Tab(11).ControlCount=   1
         Begin VB.CommandButton cmdopcion 
            Caption         =   "Agregar origen de venta a bloquear en esta caja."
            Height          =   495
            Index           =   4
            Left            =   -69120
            TabIndex        =   325
            Top             =   6480
            Width           =   5295
         End
         Begin VB.CheckBox ChkBloqueoEdicionPedidoConsolaAgr 
            Caption         =   "Bloqueo de Ediciòn de Pedidos de Consola de Agreg. (Infomatica)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   323
            Top             =   4320
            Width           =   4095
         End
         Begin VB.CheckBox ChkBloqueaPedidoSinMesa 
            Caption         =   "Bloqueo de pedidos sin mesa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   320
            Top             =   3960
            Width           =   3015
         End
         Begin VB.CheckBox ChkBloqueaCanal 
            Caption         =   "Bloqueo generar pedidos del canal 05"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   -74760
            TabIndex        =   317
            Top             =   3720
            Width           =   3975
         End
         Begin VB.CheckBox ChkBloqueaCanal 
            Caption         =   "Bloqueo generar pedidos del canal 04"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -74760
            TabIndex        =   316
            Top             =   3480
            Width           =   3975
         End
         Begin VB.CheckBox ChkBloqueaCanal 
            Caption         =   "Bloqueo generar pedidos del canal 03"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   -74760
            TabIndex        =   315
            Top             =   3240
            Width           =   3975
         End
         Begin VB.CheckBox ChkBloqueaCanal 
            Caption         =   "Bloqueo generar pedidos del canal 02"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -74760
            TabIndex        =   314
            Top             =   3000
            Width           =   3975
         End
         Begin VB.CheckBox ChkBloqueaCanal 
            Caption         =   "Bloqueo generar pedidos del canal 01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -74760
            TabIndex        =   313
            Top             =   2760
            Width           =   3975
         End
         Begin VB.CheckBox ChkBloqueoAnulaPedido 
            Caption         =   "Bloqueo de anulacion de pedidos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   311
            Top             =   2040
            Width           =   2775
         End
         Begin VB.CheckBox ChkBloqueoAnulaItemsPedido 
            Caption         =   "Bloqueo de anulacion de Items de pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   310
            Top             =   2400
            Width           =   3975
         End
         Begin VB.Frame FrmFacArgentina 
            Caption         =   "Integracion TUSFACTURASAPP"
            Height          =   1335
            Left            =   -71520
            TabIndex        =   292
            Top             =   5280
            Visible         =   0   'False
            Width           =   6495
            Begin VB.TextBox txtApiToken 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   298
               Top             =   960
               Width           =   5295
            End
            Begin VB.TextBox txtApiKey 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   297
               Top             =   600
               Width           =   5295
            End
            Begin VB.TextBox txtUserToken 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   296
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "ApiToken:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   94
               Left            =   240
               TabIndex        =   295
               Top             =   1005
               Width           =   735
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "ApiKey:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   93
               Left            =   480
               TabIndex        =   294
               Top             =   675
               Width           =   540
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "User Token:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   91
               Left            =   120
               TabIndex        =   293
               Top             =   315
               Width           =   885
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Guia"
            Height          =   975
            Left            =   -69330
            TabIndex        =   288
            Top             =   6000
            Width           =   4290
            Begin VB.CheckBox chkGuiaF2 
               Caption         =   "Imprimir Guia Formato 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   289
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "+ Activaciones"
            Height          =   6015
            Left            =   -74760
            TabIndex        =   270
            Top             =   1080
            Width           =   11055
            Begin VB.CheckBox chklActivarValesConsumo 
               Caption         =   "Activar Vales de Consumo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6480
               TabIndex        =   330
               Top             =   360
               Width           =   2760
            End
            Begin VB.CheckBox chklBusquedaReserva 
               Caption         =   "Activar Busqueda de Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   328
               Top             =   4320
               Width           =   4440
            End
            Begin VB.CheckBox chkSolicitaClaveEgreso 
               Caption         =   "Activa solicitar autorizaciòn en generaciòn de recibos de egreso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   322
               Top             =   3840
               Width           =   4500
            End
            Begin VB.CheckBox chkAutorizacionSinMesa 
               Caption         =   "Solicitar autorización para uso de la opcón ""Sin Mesa"""
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   321
               Top             =   3600
               Width           =   4500
            End
            Begin VB.CheckBox ChklActivaRestriClienteDelivery 
               Caption         =   "Activar Restricciones de Cliente Delivery"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   308
               Top             =   2880
               Width           =   4500
            End
            Begin VB.Frame Frame21 
               Caption         =   "Activaciones Mobile InfoRest"
               Height          =   1215
               Left            =   120
               TabIndex        =   304
               Top             =   4680
               Width           =   3495
               Begin VB.CheckBox chklActivarImpresionBlu 
                  Caption         =   "Activar Impresiones por Bluetooth"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   307
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   2775
               End
               Begin VB.CheckBox chklModoFacturacion 
                  Caption         =   "Activar Modo Facturación"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   306
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.CheckBox chkCajaMobile 
                  Caption         =   "Caja Mobile"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   225
                  Left            =   120
                  TabIndex        =   305
                  Top             =   360
                  Width           =   1155
               End
            End
            Begin VB.CheckBox ChklActivaCreditoCoorporativo 
               Caption         =   "Sincronizar Credito Coorporativo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   302
               Top             =   2520
               Width           =   4500
            End
            Begin VB.CheckBox ChkVistaPreviaItemsCombo 
               Caption         =   "Vista Previa de Items Combo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   299
               Top             =   2160
               Width           =   4500
            End
            Begin VB.CheckBox chklOrdenarClientes 
               Caption         =   "Ordenar Clientes con Identificador OTROS, al final de la lista, busqueda de clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   291
               Top             =   1680
               Width           =   4500
            End
            Begin VB.CheckBox chklPermitirAnularMaxDias 
               Caption         =   "Anular Documento, con clave Administrador despues de haber pasado los días como maximo nnooooo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   290
               Top             =   3120
               Visible         =   0   'False
               Width           =   4500
            End
            Begin VB.CheckBox chklActivarMensajeCtsCorrientes 
               Caption         =   "Activar Mensaje si desea imprimir en cuentas corrientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   286
               Top             =   1320
               Width           =   4500
            End
            Begin VB.CheckBox ChlNoVisualizarItemCombos 
               Caption         =   "No Visualizar  Items-Combos Automaticos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   283
               Top             =   840
               Width           =   4500
            End
            Begin VB.CheckBox chklActivaTresDecimales 
               Caption         =   "Activa visualización de cantidad detalle de pedido con 3 decimales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   271
               Top             =   360
               Width           =   4500
            End
            Begin VB.Label Label11 
               Caption         =   "(Solo punto de venta y punto de adición)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   360
               TabIndex        =   284
               Top             =   1080
               Width           =   3015
            End
         End
         Begin VB.CheckBox ChkBloqAfectoInafecto 
            Caption         =   "Bloquear emisión de documentos por consumo con Items afectos e inafectos."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -74760
            TabIndex        =   262
            Top             =   1560
            Width           =   4065
         End
         Begin VB.CheckBox ChkBloqEmiDoc 
            Caption         =   "Bloquear emisión de documentos por consumo con bonificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -74760
            TabIndex        =   255
            Top             =   1080
            Width           =   4065
         End
         Begin VB.Frame Frame17 
            Caption         =   "Activaciones Integraciones de Venta"
            Height          =   3855
            Left            =   -71500
            TabIndex        =   247
            Top             =   1320
            Width           =   3375
            Begin VB.CheckBox chkIntegracionWebApi 
               Caption         =   "Integracion Web Api"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   303
               Top             =   1200
               Width           =   2025
            End
            Begin VB.CheckBox chklIntMobileKDS 
               Caption         =   "Integración Mobile - KDS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   287
               Top             =   920
               Width           =   2175
            End
            Begin VB.CheckBox chkIntegracionPatio 
               Caption         =   "Integracion Patio de Comidas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   261
               Top             =   600
               Width           =   2505
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Integracion Rappi"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   250
               Top             =   1680
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.CheckBox chkActivaUber 
               Caption         =   "Integracion Uber Eats"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   249
               Top             =   360
               Width           =   2025
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Activaciones Forma de Pagos"
            Height          =   5295
            Left            =   -74920
            TabIndex        =   246
            Top             =   1320
            Width           =   3375
            Begin VB.CheckBox chkIntFPAY 
               Caption         =   "Integracion FPAY"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   301
               Top             =   3720
               Width           =   1695
            End
            Begin VB.CheckBox chkPagoEfectivo 
               Caption         =   "Integracion PagoEfectivo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   264
               Top             =   3000
               Width           =   3105
            End
            Begin VB.CheckBox chkIntMercadoPagoQREstatico 
               Caption         =   "Integracion Mercado Pago QR estatico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   263
               Top             =   2760
               Width           =   3105
            End
            Begin VB.CheckBox chkIntMercadoPagoQR 
               Caption         =   "Integracion Mercado Pago QR Dinamico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   259
               Top             =   1920
               Width           =   3105
            End
            Begin VB.TextBox txtIpCashDrow 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   840
               TabIndex        =   257
               Text            =   "192.168.1.1"
               Top             =   1560
               Width           =   2295
            End
            Begin VB.CheckBox ChkActivaCashDro 
               Caption         =   "Integracion CashDrow"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   254
               Top             =   1080
               Width           =   3105
            End
            Begin VB.CheckBox chkIntVisaNetQR 
               Caption         =   "Integracion VisaNet QR"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   248
               Top             =   360
               Width           =   3105
            End
            Begin VB.Label Label 
               Caption         =   "(Aplica a formulario de pago de documentos)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Index           =   43
               Left            =   360
               TabIndex        =   265
               Top             =   3360
               Width           =   2895
            End
            Begin VB.Label Label 
               Caption         =   "(Aplica solo Punto Venta, Caja Rapida y Generacion de documentos)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   40
               Left            =   360
               TabIndex        =   260
               Top             =   2400
               Width           =   2895
            End
            Begin VB.Label lblIpCashDrow 
               Caption         =   "IP:"
               Height          =   255
               Left            =   480
               TabIndex        =   258
               Top             =   1560
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "(Aplica CashDrow3 y CashDrow5)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   41
               Left            =   360
               TabIndex        =   256
               Top             =   1320
               Width           =   2775
            End
            Begin VB.Label Label 
               Caption         =   "(Aplica solo Punto Venta, Caja Rapida y Generacion de documentos)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   42
               Left            =   360
               TabIndex        =   253
               Top             =   600
               Width           =   2895
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Comandas"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   241
            Top             =   5880
            Width           =   5460
            Begin VB.CheckBox ChkImpComandaBarra 
               Caption         =   "Imprimir Comanda por Item con Codigo QR"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   251
               Top             =   480
               Width           =   4155
            End
            Begin VB.CheckBox chkComandaF2 
               Caption         =   "Imprimir Comanda Formato 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   242
               Top             =   240
               Width           =   2595
            End
            Begin VB.Label Label 
               Caption         =   "(No Aplica a Adicion)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Index           =   44
               Left            =   480
               TabIndex        =   252
               Top             =   720
               Width           =   3135
            End
         End
         Begin VB.CheckBox chkConsumo4 
            Caption         =   "Emisión por consumo en Cuenta Corriente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   191
            Top             =   4180
            Width           =   3570
         End
         Begin VB.CommandButton cmdQuitarFotoPie 
            Caption         =   "Quitar Imagen"
            Height          =   350
            Left            =   -69480
            TabIndex        =   184
            Top             =   4980
            Width           =   4000
         End
         Begin VB.CommandButton cmdQuitarFotoCabecera 
            Caption         =   "Quitar Imagen"
            Height          =   350
            Left            =   -74520
            TabIndex        =   183
            Top             =   4980
            Width           =   4000
         End
         Begin VB.CommandButton cmdAgregarFotoPie 
            Caption         =   "Editar Imagen Pie"
            Height          =   350
            Left            =   -69480
            TabIndex        =   182
            Top             =   4620
            Width           =   4000
         End
         Begin VB.CommandButton cmdAgregarFoto 
            Caption         =   "Editar Imagen Cabecera"
            Height          =   350
            Left            =   -74520
            TabIndex        =   181
            Top             =   4620
            Width           =   4000
         End
         Begin VB.Frame fraAreaChef 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3675
            Left            =   -74880
            TabIndex        =   174
            Top             =   1620
            Width           =   9830
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   10
               Left            =   1515
               Picture         =   "frmCajaDetalle.frx":0592
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   1650
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   11
               Left            =   2775
               Picture         =   "frmCajaDetalle.frx":0AC4
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   1650
               Width           =   1215
            End
            Begin VB.CheckBox chkAreaChef 
               Alignment       =   1  'Right Justify
               Caption         =   "Area Chef :  "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   360
               TabIndex        =   175
               Top             =   1005
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo cboAreaChef 
               Height          =   315
               Left            =   1305
               TabIndex        =   178
               Top             =   465
               Width           =   2970
               _ExtentX        =   5239
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Area :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   735
               TabIndex        =   179
               Top             =   525
               Width           =   420
            End
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   12
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":0BC6
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   3180
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   13
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":0CC8
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   1620
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   14
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":11FA
            Style           =   1  'Graphical
            TabIndex        =   170
            Top             =   2385
            Width           =   1215
         End
         Begin VB.CheckBox chkMulti2 
            Caption         =   "Areas de Producción por Sub Grupos"
            Height          =   255
            Left            =   -74640
            TabIndex        =   157
            Top             =   2220
            Width           =   3855
         End
         Begin VB.Frame fra2 
            Enabled         =   0   'False
            Height          =   2775
            Left            =   -74760
            TabIndex        =   158
            Top             =   2220
            Width           =   9375
            Begin VB.Frame fraAreaProduccion 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2460
               Left            =   120
               TabIndex        =   159
               Top             =   240
               Width           =   9120
               Begin VB.CommandButton cmdOpcionGrilla 
                  Caption         =   "Cancelar"
                  Height          =   645
                  Index           =   16
                  Left            =   6120
                  Picture         =   "frmCajaDetalle.frx":172C
                  Style           =   1  'Graphical
                  TabIndex        =   161
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.CommandButton cmdOpcionGrilla 
                  Caption         =   "Grabar"
                  Height          =   645
                  Index           =   15
                  Left            =   4920
                  Picture         =   "frmCajaDetalle.frx":182E
                  Style           =   1  'Graphical
                  TabIndex        =   160
                  Top             =   600
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo cboSubGrupo 
                  Height          =   315
                  Left            =   1665
                  TabIndex        =   162
                  Top             =   480
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   556
                  _Version        =   393216
                  MatchEntry      =   -1  'True
                  Style           =   2
                  BackColor       =   16777215
                  ListField       =   ""
                  BoundColumn     =   ""
                  Text            =   ""
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
               Begin MSDataListLib.DataCombo cboAreaProd 
                  Height          =   315
                  Left            =   1665
                  TabIndex        =   163
                  Top             =   945
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   556
                  _Version        =   393216
                  MatchEntry      =   -1  'True
                  Style           =   2
                  BackColor       =   16777215
                  ListField       =   ""
                  BoundColumn     =   ""
                  Text            =   ""
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
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Area Producción :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   23
                  Left            =   255
                  TabIndex        =   165
                  Top             =   1005
                  Width           =   1275
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Sub Grupo :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   22
                  Left            =   540
                  TabIndex        =   164
                  Top             =   525
                  Width           =   855
               End
            End
            Begin TrueOleDBGrid80.TDBGrid grdGrillaSubgrupos 
               Height          =   2295
               Left            =   120
               TabIndex        =   166
               Top             =   360
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   4048
               _LayoutType     =   4
               _RowHeight      =   23
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   ""
               Columns(0).NumberFormat=   "True/False"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   1
               Splits(0)._UserFlags=   0
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0).ScrollBars=   2
               Splits(0).AllowColSelect=   0   'False
               Splits(0).FetchRowStyle=   -1  'True
               Splits(0).DividerStyle=   2
               Splits(0).DividerColor=   32768
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=1"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   12632256
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               DirectionAfterTab=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
               _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(9)   =   ":id=4,.fontname=Arial"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
               _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=Arial"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
               _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
               _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
               _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
               _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
               _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
               _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
               _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
               _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
               _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
               _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
               _StyleDefs(38)  =   "Named:id=33:Normal"
               _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
               _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
               _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
               _StyleDefs(43)  =   "Named:id=34:Heading"
               _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
               _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
               _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(47)  =   ":id=34,.fontname=Arial"
               _StyleDefs(48)  =   "Named:id=35:Footing"
               _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(50)  =   "Named:id=36:Selected"
               _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
               _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(53)  =   ":id=36,.fontname=Arial"
               _StyleDefs(54)  =   "Named:id=37:Caption"
               _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(56)  =   "Named:id=38:HighlightRow"
               _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
               _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
               _StyleDefs(60)  =   "Named:id=39:EvenRow"
               _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(62)  =   "Named:id=40:OddRow"
               _StyleDefs(63)  =   ":id=40,.parent=33"
               _StyleDefs(64)  =   "Named:id=41:RecordSelector"
               _StyleDefs(65)  =   ":id=41,.parent=34"
               _StyleDefs(66)  =   "Named:id=42:FilterBar"
               _StyleDefs(67)  =   ":id=42,.parent=33"
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Modificar"
               Height          =   645
               Index           =   19
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":1D60
               Style           =   1  'Graphical
               TabIndex        =   167
               Top             =   1200
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Agregar"
               Height          =   645
               Index           =   18
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":2292
               Style           =   1  'Graphical
               TabIndex        =   168
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Eliminar"
               Height          =   645
               Index           =   17
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":27C4
               Style           =   1  'Graphical
               TabIndex        =   169
               Top             =   1920
               Width           =   1215
            End
         End
         Begin VB.CheckBox chkMulti1 
            Caption         =   "Una sola Area de Producción"
            Height          =   375
            Left            =   -74640
            TabIndex        =   153
            Top             =   1260
            Width           =   3015
         End
         Begin VB.Frame fra1 
            Enabled         =   0   'False
            Height          =   735
            Left            =   -74760
            TabIndex        =   154
            Top             =   1380
            Width           =   9375
            Begin MSDataListLib.DataCombo cboAreaProduccion 
               Height          =   315
               Left            =   1740
               TabIndex        =   155
               Top             =   360
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Área de Producción :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   156
               Top             =   420
               Width           =   1500
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   " Enlace SIAB "
            Height          =   735
            Left            =   -69285
            TabIndex        =   138
            Top             =   2820
            Width           =   4185
            Begin VB.CheckBox chkSiab 
               Caption         =   "Caja conectada al POS SIAB"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   225
               TabIndex        =   139
               Top             =   375
               Width           =   3825
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " Otros Documentos "
            Height          =   1335
            Left            =   -69330
            TabIndex        =   134
            Top             =   4680
            Width           =   4290
            Begin VB.CheckBox chkCodigoReciboIngreso 
               Caption         =   "Impresión Código de Barras en Recibo de Ingreso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   188
               Top             =   720
               Visible         =   0   'False
               Width           =   4020
            End
            Begin VB.TextBox txtLimiteReimpresion 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               TabIndex        =   147
               Text            =   "0"
               Top             =   945
               Width           =   600
            End
            Begin VB.CheckBox chkValor 
               Caption         =   "Impresión Valorizada de las Cortesías"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   137
               Top             =   480
               Width           =   3570
            End
            Begin VB.CheckBox chkCambioMesa 
               Caption         =   "Impresion de Cambio de Mesa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   136
               Top             =   240
               Width           =   3570
            End
            Begin VB.Label Label4 
               Caption         =   "Permite"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   149
               Top             =   990
               Width           =   555
            End
            Begin VB.Label Label5 
               Caption         =   "reimpresiones por Pedido (0 sin limite)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1410
               TabIndex        =   148
               Top             =   990
               Width           =   2625
            End
         End
         Begin VB.Frame fraArea 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3915
            Left            =   -74865
            TabIndex        =   48
            Top             =   1320
            Width           =   9855
            Begin VB.CheckBox lActivaDobleImpComanda 
               Alignment       =   1  'Right Justify
               Caption         =   "Activar doble Impresion de Comandas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1080
               TabIndex        =   331
               Top             =   1560
               Width           =   3180
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   8
               Left            =   7155
               Picture         =   "frmCajaDetalle.frx":28C6
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   2970
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   9
               Left            =   8415
               Picture         =   "frmCajaDetalle.frx":2DF8
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   2970
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo cboArea 
               Height          =   315
               Left            =   1305
               TabIndex        =   51
               Top             =   465
               Width           =   2970
               _ExtentX        =   5239
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin MSDataListLib.DataCombo cboImpArea 
               Height          =   315
               Left            =   1305
               TabIndex        =   52
               Top             =   945
               Width           =   2970
               _ExtentX        =   5239
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Area :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   735
               TabIndex        =   54
               Top             =   525
               Width           =   420
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresora :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   375
               TabIndex        =   53
               Top             =   1005
               Width           =   780
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   " Documentos "
            Height          =   3540
            Left            =   -69330
            TabIndex        =   118
            Top             =   1080
            Width           =   4290
            Begin VB.CheckBox chkimprimePropinaDoc 
               Caption         =   "Mostrar propina en Documento electronico (Proceso registro de pago antes de impresion)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   300
               Top             =   2880
               Width           =   4095
            End
            Begin VB.CheckBox ChlMostrarPrecioItemsCombo 
               Caption         =   "Mostrar Precio de Items-Combo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   285
               Top             =   2640
               Width           =   3855
            End
            Begin VB.CheckBox chkImpClienteCab 
               Caption         =   "Impresión de cliente cabecera de Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   244
               Top             =   2400
               Width           =   3930
            End
            Begin VB.CheckBox chkMotDesc 
               Caption         =   "Impresión de Motivo de Descuento en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   238
               Top             =   2160
               Width           =   4050
            End
            Begin VB.CheckBox chkObservacionCabDoc 
               Caption         =   "Impresión de Observacion en cabecera del Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   212
               Top             =   1560
               Width           =   3930
            End
            Begin VB.CheckBox chkDescripcionAlternativa 
               Caption         =   "Impresión Descripción Alternativa en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   130
               Top             =   1920
               Width           =   3930
            End
            Begin VB.CheckBox chkPropiedadDocumento 
               Caption         =   "Impresión de Propiedades en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   129
               Top             =   840
               Width           =   3870
            End
            Begin VB.CheckBox chkComboDocumento 
               Caption         =   "Impresion de Combos en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   128
               Top             =   600
               Width           =   3570
            End
            Begin VB.CheckBox chkDocumentoAgrupado 
               Caption         =   "Impresion Agrupada de Items de Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   127
               Top             =   315
               Width           =   3615
            End
            Begin VB.CheckBox chkObservacionDocumento 
               Caption         =   "Impresión de Observaciones en Documentos (Combos y Propiedades)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   120
               TabIndex        =   126
               Top             =   1080
               Width           =   3810
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " Precuenta "
            Height          =   4740
            Left            =   -74865
            TabIndex        =   114
            Top             =   1080
            Width           =   5460
            Begin VB.CheckBox chkPrecuentaImprime 
               Caption         =   "Impresión de forma de pago en precuenta."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   327
               Top             =   3960
               Width           =   4935
            End
            Begin VB.CheckBox chkPrecuentaCabUnica 
               Caption         =   "Impresión de Precuenta con Cabecera Única"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   319
               Top             =   3720
               Width           =   4815
            End
            Begin VB.CheckBox chkOcultarDesgloseRazonSocial 
               Caption         =   "Ocultar desglose de de razon social, ruc y correo en precuenta."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   312
               Top             =   3480
               Width           =   5055
            End
            Begin VB.CheckBox chkImprimePropinaSugerida 
               Caption         =   "Impresión de propina sugerida en precuenta."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   268
               Top             =   3240
               Width           =   3615
            End
            Begin VB.CheckBox chkImpPropina 
               Caption         =   "Solicitar Propina en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   240
               Top             =   3000
               Width           =   3570
            End
            Begin VB.CheckBox chkPrecuentaNoValorizada 
               Caption         =   "Impresión de Precuenta no valorizada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   192
               Top             =   2760
               Width           =   3570
            End
            Begin VB.CheckBox chkImprimeImagCabPrecuenta 
               Caption         =   "Impresión Imagen en Cabecera de Documento en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   123
               Top             =   2040
               Width           =   4995
            End
            Begin VB.CheckBox chkImprimeImagPiePrecuenta 
               Caption         =   "Impresión Imagen en Pie de Documento en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   124
               Top             =   2280
               Width           =   4395
            End
            Begin VB.CheckBox chkBloqueaPrecuenta 
               Caption         =   "No Permitir Emisión de Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   125
               Top             =   2520
               Width           =   3570
            End
            Begin VB.TextBox txtLimitePrecuenta 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               TabIndex        =   144
               Text            =   "0"
               Top             =   4335
               Width           =   600
            End
            Begin VB.CheckBox chkEquivaPrecuenta 
               Caption         =   "Impresión Equivalencia Dolares en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   122
               Top             =   1800
               Width           =   3570
            End
            Begin VB.CheckBox chkPrecioNetoPrecuenta 
               Caption         =   "Impresión de Precio Neto en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   120
               Top             =   1320
               Width           =   3570
            End
            Begin VB.CheckBox chkImpuestoPrecuenta 
               Caption         =   "Impresion Impuestos desglos. en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   121
               Top             =   1560
               Width           =   3570
            End
            Begin VB.CheckBox chkPropiedadPrecuenta 
               Caption         =   "Impresión de Propiedades  en Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   117
               Top             =   820
               Width           =   3990
            End
            Begin VB.CheckBox chkAgrupada 
               Caption         =   "Impresion Agrupada de Items de Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   115
               Top             =   315
               Width           =   3480
            End
            Begin VB.CheckBox chkComboPrecuenta 
               Caption         =   "Impresion de Combos en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   116
               Top             =   560
               Width           =   3570
            End
            Begin VB.CheckBox chkObservacionPrecuenta 
               Caption         =   "Impresión de Observaciones en Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   119
               Top             =   1080
               Width           =   3990
            End
            Begin VB.Label Label2 
               Caption         =   "precuentas por Pedido (0 sin limite)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   146
               Top             =   4380
               Width           =   2580
            End
            Begin VB.Label Label3 
               Caption         =   "Permite"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   145
               Top             =   4380
               Width           =   555
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " Activaciones Básicas "
            Height          =   6060
            Left            =   -74775
            TabIndex        =   96
            Top             =   975
            Width           =   5010
            Begin VB.CheckBox chkMotivoAdm 
               Caption         =   "Pide un Motivo Administrativo de NC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   135
               TabIndex        =   309
               Top             =   5470
               Width           =   4440
            End
            Begin VB.CheckBox chklActivarAvisoRetenciones 
               Caption         =   "Activar Aviso para Clientes de Retenciòn"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   135
               TabIndex        =   269
               Top             =   5160
               Width           =   4455
            End
            Begin VB.CheckBox chkBuscarPedidoBarra 
               Caption         =   "Permite buscar pedidos facturados y no enviados por codigo de Barra"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   135
               TabIndex        =   245
               ToolTipText     =   "Activado: Pedidos con Mesa/ Desactivado: Pedidos con Cliente"
               Top             =   4800
               Width           =   4770
            End
            Begin VB.CheckBox chkClaveEnvio 
               Caption         =   "Solicitar clave para envio a Producción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   201
               ToolTipText     =   "Activado: Pedidos con Mesa/ Desactivado: Pedidos con Cliente"
               Top             =   4540
               Width           =   3810
            End
            Begin VB.CheckBox chkBuscaPedidoFiltrarMesa 
               Caption         =   "Buscar Pedido: Filtrar Pedidos con Mesa asignada en Grilla"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   195
               ToolTipText     =   "Activado: Pedidos con Mesa/ Desactivado: Pedidos con Cliente"
               Top             =   4285
               Width           =   4770
            End
            Begin VB.CheckBox chkBuscaPedidoVisualizaGrilla 
               Caption         =   "Buscar Pedido: Ingreso directo a Visualización en Grilla"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   194
               ToolTipText     =   "Activado: Acceso a visualizar grilla / Desactivado: Acceso a visualizar mapa de mesas"
               Top             =   4035
               Width           =   4770
            End
            Begin VB.CheckBox chkPagoRapidoMod 
               Caption         =   "Ingreso a Pago Rápido desde Modifica Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   135
               TabIndex        =   189
               Top             =   3350
               Width           =   3735
            End
            Begin VB.CheckBox chkBuscaPedido 
               Caption         =   "Búsqueda Predeterminada por Número de Pedido al Transferir"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   103
               ToolTipText     =   "Activado: Por Número Pedido / Desactivado: Por Mesa"
               Top             =   1800
               Width           =   4770
            End
            Begin VB.CheckBox chkAccesoDespachoPedido 
               Caption         =   "Activa Acceso a Despachos en Central de Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   186
               Top             =   3800
               Width           =   4770
            End
            Begin VB.CheckBox chkHuella 
               Caption         =   "Activa Ingreso Directo a Huella Digital"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   180
               Top             =   3590
               Width           =   3570
            End
            Begin VB.CheckBox chkCajaRapida 
               Caption         =   "Ingreso Directo a la Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   133
               Top             =   2680
               Width           =   3570
            End
            Begin VB.CheckBox chkPagoRapido 
               Caption         =   "Ingreso a Pago Rápido desde Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   132
               Top             =   2900
               Width           =   3570
            End
            Begin VB.CheckBox chkPagoRapidopv 
               Caption         =   "Ingreso a Pago Rápido desde Punto de Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   131
               Top             =   3140
               Width           =   3570
            End
            Begin VB.CheckBox chkPreCuenta 
               Caption         =   "Activa el poder cambiar de impresora de PreCuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   113
               Top             =   2265
               Width           =   4560
            End
            Begin VB.CheckBox chkDisgrega 
               Caption         =   "Activa el disgregar el Producto en dos partes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   112
               Top             =   2480
               Width           =   3930
            End
            Begin VB.CheckBox chkFiltroTipoPedido 
               Caption         =   "Solo importar pedidos del canal de venta 01 (no permite otros canales de venta)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   325
               Left            =   135
               TabIndex        =   104
               Top             =   1240
               Width           =   4800
            End
            Begin VB.CheckBox chkAdicion 
               Caption         =   "Activa las Transferencias (Importar Pedidos)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   102
               Top             =   1600
               Width           =   3570
            End
            Begin VB.CheckBox chkModificaTipoPedido 
               Caption         =   "Activa el poder Modificar el Tipo de Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   105
               Top             =   2025
               Width           =   3705
            End
            Begin VB.CheckBox chkOrden 
               Caption         =   "Mantiene el Orden 1 de todos los SubGrupos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   101
               Top             =   195
               Width           =   3570
            End
            Begin VB.CheckBox chkElimina 
               Caption         =   "Pide un Motivo de Eliminación en el Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   100
               Top             =   1030
               Width           =   3570
            End
            Begin VB.CheckBox chkEliminaC 
               Caption         =   "Pide un Motivo de Eliminación en el Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   99
               Top             =   830
               Width           =   3570
            End
            Begin VB.CheckBox chkDirecto 
               Caption         =   "Activa el Control de Envíos Directos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   98
               Top             =   400
               Width           =   3570
            End
            Begin VB.CheckBox chkVComanda 
               Caption         =   "Activa el ingreso de Comandas Manuales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   97
               Top             =   610
               Width           =   3570
            End
         End
         Begin VB.Frame fraGrilla 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4995
            Left            =   -74865
            TabIndex        =   34
            Top             =   1260
            Width           =   9825
            Begin VB.CheckBox chkSolicitarConfirmacion 
               Alignment       =   1  'Right Justify
               Caption         =   "Solicitar confirmación para generar documentos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   318
               Top             =   2640
               Width           =   5055
            End
            Begin VB.CheckBox chkMayorCero 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresion de platos con monto mayor ""0"":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7680
               TabIndex        =   237
               Top             =   1560
               Width           =   2055
            End
            Begin VB.CheckBox chkCodProdDes 
               Alignment       =   1  'Right Justify
               Caption         =   "Impr. CodigoProducto, Descuento/unitario:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7680
               TabIndex        =   236
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Frame Frame14 
               Caption         =   "Desglose en Impresión :"
               Height          =   1815
               Left            =   5280
               TabIndex        =   213
               Top             =   240
               Width           =   2295
               Begin VB.CheckBox chkImpuesto1 
                  Caption         =   "Check1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   217
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkImpuesto2 
                  Caption         =   "Check2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   216
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkImpuesto3 
                  Caption         =   "Check3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   215
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkopGravInaf 
                  Caption         =   "Op. Gravada, Inafecta, Exoneración"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   214
                  Top             =   960
                  Width           =   1965
               End
            End
            Begin VB.CheckBox chkImpResumido 
               Alignment       =   1  'Right Justify
               Caption         =   "Impr. Detallado:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7680
               TabIndex        =   211
               Top             =   760
               Width           =   2055
            End
            Begin VB.TextBox txtPrefijoEnlace 
               Height          =   285
               Left            =   8880
               MaxLength       =   1
               TabIndex        =   203
               Top             =   2235
               Width           =   795
            End
            Begin VB.TextBox txtAutorizacion 
               Height          =   285
               Left            =   2040
               MaxLength       =   20
               TabIndex        =   196
               Top             =   3120
               Width           =   3120
            End
            Begin VB.CheckBox chkFacturacionOfisis 
               Alignment       =   1  'Right Justify
               Caption         =   "Documento Electrónico Ofisis :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5520
               TabIndex        =   193
               Top             =   2880
               Width           =   2565
            End
            Begin VB.CheckBox chkFacturacionE 
               Alignment       =   1  'Right Justify
               Caption         =   "Facturación Electrónica:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5520
               TabIndex        =   187
               Top             =   2280
               Width           =   2055
            End
            Begin VB.CheckBox chLImprimeImageCab 
               Alignment       =   1  'Right Justify
               Caption         =   "En Cabecera"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               TabIndex        =   6
               Top             =   2280
               Width           =   1335
            End
            Begin VB.CheckBox chLImprimeImagePie 
               Alignment       =   1  'Right Justify
               Caption         =   "En Pie"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4080
               TabIndex        =   7
               Top             =   2280
               Width           =   1095
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   4
               Left            =   8370
               Picture         =   "frmCajaDetalle.frx":2EFA
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   4230
               Width           =   1215
            End
            Begin VB.TextBox txtSerie 
               Height          =   285
               Left            =   2040
               MaxLength       =   5
               TabIndex        =   1
               Top             =   735
               Width           =   1410
            End
            Begin VB.TextBox txtDescripcion 
               Height          =   285
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   5
               Top             =   1815
               Width           =   3120
            End
            Begin VB.CheckBox chkResumen 
               Alignment       =   1  'Right Justify
               Caption         =   "Resumen :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   8
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   3
               Left            =   7065
               Picture         =   "frmCajaDetalle.frx":2FFC
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   4230
               Width           =   1215
            End
            Begin VB.CheckBox chkDocEquivDolares 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresión Equi. Dolares"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7680
               TabIndex        =   9
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtCorrelativo 
               Height          =   285
               Left            =   3525
               MaxLength       =   9
               TabIndex        =   2
               Top             =   735
               Width           =   1635
            End
            Begin VB.TextBox txtSerie2 
               Height          =   285
               Left            =   3315
               MaxLength       =   3
               TabIndex        =   37
               Text            =   "000"
               Top             =   735
               Width           =   645
            End
            Begin VB.TextBox txtCorrelativo2 
               Height          =   285
               Left            =   3900
               MaxLength       =   9
               TabIndex        =   35
               Text            =   "00000000"
               Top             =   735
               Width           =   1140
            End
            Begin MSDataListLib.DataCombo cboLocal 
               Height          =   315
               Left            =   2040
               TabIndex        =   36
               Top             =   735
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin MSDataListLib.DataCombo cboTipoDocumento 
               Height          =   315
               Left            =   2040
               TabIndex        =   0
               Top             =   315
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin MSDataListLib.DataCombo cboImpresora 
               Height          =   315
               Left            =   2040
               TabIndex        =   3
               Top             =   1065
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin MSDataListLib.DataCombo cboFormulario 
               Height          =   315
               Left            =   2040
               TabIndex        =   4
               Top             =   1440
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
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
            Begin MSComCtl2.DTPicker dtpFechaInicio 
               Height          =   345
               Left            =   2040
               TabIndex        =   199
               Top             =   3480
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   50593795
               CurrentDate     =   37795
            End
            Begin MSComCtl2.DTPicker dtpFechaCaducida 
               Height          =   345
               Left            =   2040
               TabIndex        =   200
               Top             =   3960
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   50593795
               CurrentDate     =   37795
            End
            Begin VB.Frame Frame13 
               Caption         =   "(solo aplica en el enlace con OFISIS)"
               ForeColor       =   &H00800000&
               Height          =   1455
               Left            =   5400
               TabIndex        =   205
               Top             =   2640
               Width           =   4335
               Begin VB.TextBox txtCompVenta 
                  Height          =   285
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   209
                  Top             =   960
                  Width           =   2040
               End
               Begin VB.TextBox txtFormVenta 
                  Height          =   285
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   208
                  Top             =   555
                  Width           =   2040
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Comprobante Venta :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   30
                  Left            =   120
                  TabIndex        =   207
                  Top             =   1005
                  Width           =   1875
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Formulario Venta :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   29
                  Left            =   360
                  TabIndex        =   206
                  Top             =   600
                  Width           =   1650
               End
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Prefijo Enlace:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   7800
               TabIndex        =   202
               Top             =   2280
               Width           =   1020
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Caducidad :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   27
               Left            =   600
               TabIndex        =   198
               Top             =   4035
               Width           =   1350
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Inicio :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   26
               Left            =   960
               TabIndex        =   197
               Top             =   3555
               Width           =   960
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresión de Imágenes  :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   165
               TabIndex        =   185
               Top             =   2280
               Width           =   1770
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   660
               TabIndex        =   43
               Top             =   375
               Width           =   1275
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Serie y Correlativo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   570
               TabIndex        =   42
               Top             =   825
               Width           =   1365
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresora :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   1155
               TabIndex        =   41
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Formulario :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   1125
               TabIndex        =   40
               Top             =   1500
               Width           =   810
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Descripción :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   1005
               TabIndex        =   39
               Top             =   1860
               Width           =   930
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Autorización :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   975
               TabIndex        =   38
               Top             =   3165
               Width           =   960
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Lectora de Barras  "
            Height          =   1395
            Left            =   -69285
            TabIndex        =   73
            Top             =   1365
            Width           =   4185
            Begin VB.CheckBox chkEAN13 
               Caption         =   "EAN13"
               Height          =   195
               Left            =   2280
               TabIndex        =   210
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton opcCapturaPeso 
               Caption         =   "Captura el Peso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2190
               TabIndex        =   152
               Top             =   1040
               Width           =   1575
            End
            Begin VB.OptionButton opcCapturaPrecio 
               Caption         =   "Captura el Precio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   270
               TabIndex        =   151
               Top             =   1040
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.CheckBox chkRotulado 
               Caption         =   "Enlace Rotulado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   270
               TabIndex        =   150
               Top             =   720
               Width           =   1665
            End
            Begin VB.TextBox txtLongitudBarra 
               Height          =   330
               Left            =   2250
               TabIndex        =   74
               Top             =   307
               Width           =   870
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Longitud Código Barras"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   270
               TabIndex        =   75
               Top             =   375
               Width           =   1650
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Enlace VisaNet "
            Height          =   780
            Left            =   -69285
            TabIndex        =   71
            Top             =   3660
            Visible         =   0   'False
            Width           =   4185
            Begin VB.CheckBox chkVisaNet 
               Caption         =   "Caja conectada al POS Pin Pad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   225
               TabIndex        =   72
               Top             =   375
               Width           =   3825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Visor de Precios "
            Height          =   1695
            Left            =   -74850
            TabIndex        =   64
            Top             =   1365
            Width           =   5430
            Begin VB.TextBox txtMensaje2 
               Height          =   285
               Left            =   1695
               MaxLength       =   19
               TabIndex        =   67
               Top             =   1200
               Width           =   3585
            End
            Begin VB.TextBox txtMensaje1 
               Height          =   285
               Left            =   1695
               MaxLength       =   19
               TabIndex        =   66
               Top             =   840
               Width           =   3585
            End
            Begin VB.TextBox txtPuerto 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1695
               MaxLength       =   1
               TabIndex        =   65
               Top             =   420
               Width           =   735
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Mensaje 2 :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   705
               TabIndex        =   70
               Top             =   1245
               Width           =   825
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Mensaje 1 :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   705
               TabIndex        =   69
               Top             =   885
               Width           =   825
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Puerto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   975
               TabIndex        =   68
               Top             =   465
               Width           =   555
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " Balanza Electrónica "
            Height          =   2535
            Left            =   -74880
            TabIndex        =   61
            Top             =   3060
            Width           =   5385
            Begin VB.CheckBox chklIntegracionCASSWII30CD 
               Caption         =   "Integracion con balanza CAS SWII - 30CD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   272
               Top             =   2160
               Width           =   3615
            End
            Begin VB.TextBox txtbaltiempo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   2
               TabIndex        =   234
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtBalcomando 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   233
               Top             =   600
               Width           =   615
            End
            Begin VB.CommandButton cmdGuardarBalanza 
               Caption         =   "Guardar"
               Height          =   375
               Left            =   4440
               TabIndex        =   231
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox cboBal2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":352E
               Left            =   1440
               List            =   "frmCajaDetalle.frx":3541
               TabIndex        =   230
               Top             =   600
               Width           =   2055
            End
            Begin VB.ComboBox cboBal5 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3559
               Left            =   1440
               List            =   "frmCajaDetalle.frx":3566
               TabIndex        =   229
               Top             =   1680
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.ComboBox cboBal3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3589
               Left            =   1440
               List            =   "frmCajaDetalle.frx":359C
               TabIndex        =   228
               Top             =   960
               Width           =   2055
            End
            Begin VB.ComboBox cboBal4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":35C5
               Left            =   1440
               List            =   "frmCajaDetalle.frx":35D2
               TabIndex        =   227
               Top             =   1320
               Width           =   2055
            End
            Begin VB.ComboBox cboBal1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":35E5
               Left            =   1440
               List            =   "frmCajaDetalle.frx":361F
               TabIndex        =   226
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox chkActivoBal 
               Alignment       =   1  'Right Justify
               Caption         =   "Activo"
               Height          =   195
               Left            =   4080
               TabIndex        =   220
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtBalanzaPuerto 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   1
               TabIndex        =   62
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Tiempo Espera Bal.(Seg):"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   37
               Left            =   3480
               TabIndex        =   235
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Comando:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   3720
               TabIndex        =   232
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Control de flujo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   120
               TabIndex        =   225
               Top             =   1800
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Bits de parada :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   240
               TabIndex        =   224
               Top             =   1440
               Width           =   1110
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Paridad :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   33
               Left            =   720
               TabIndex        =   223
               Top             =   1080
               Width           =   630
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Bits de datos :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   32
               Left            =   360
               TabIndex        =   222
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Bits por segundo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   31
               Left            =   120
               TabIndex        =   221
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Serial :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   3600
               TabIndex        =   63
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Texto Predeterminado Por Consumo "
            Height          =   825
            Left            =   -74880
            TabIndex        =   59
            Top             =   5760
            Width           =   9705
            Begin VB.TextBox txtTextoConsumo 
               Height          =   285
               Left            =   180
               MaxLength       =   200
               TabIndex        =   60
               Top             =   360
               Width           =   9345
            End
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   7
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":3690
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   3030
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   5
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":3792
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1500
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   6
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":3CC4
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2265
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   2
            Left            =   -66030
            Picture         =   "frmCajaDetalle.frx":41F6
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   4500
            Width           =   975
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   1
            Left            =   -67080
            Picture         =   "frmCajaDetalle.frx":42F8
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   4500
            Width           =   975
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   0
            Left            =   -68160
            Picture         =   "frmCajaDetalle.frx":482A
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4500
            Width           =   975
         End
         Begin VB.CheckBox chkConsumo3 
            Caption         =   "Emisión por consumo en Caja Rápida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   33
            Top             =   3870
            Width           =   3570
         End
         Begin VB.CheckBox chkConsumo2 
            Caption         =   "Emisión por consumo en Pagos y División "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   32
            Top             =   3555
            Width           =   3570
         End
         Begin VB.CheckBox chkConsumo1 
            Caption         =   "Emisión por consumo en Emisión Rápida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   31
            Top             =   3225
            Width           =   3570
         End
         Begin VB.CheckBox chkObservacion 
            Caption         =   "Obligatoriedad de Observación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   30
            Top             =   2925
            Width           =   3570
         End
         Begin VB.CheckBox chkObligaPrecuenta 
            Caption         =   "Obligatoriedad de Impresión de Precuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   28
            Top             =   2235
            Width           =   3570
         End
         Begin VB.CheckBox chkCancelacion 
            Caption         =   "Obligatoriedad de Cancelación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   29
            Top             =   2565
            Width           =   3570
         End
         Begin VB.CheckBox chkObligaPrinter 
            Caption         =   "Obligatoriedad de Impresión de Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   27
            Top             =   1890
            Width           =   3570
         End
         Begin VB.CheckBox chkComanda 
            Caption         =   "Obligatoriedad de Comandas Manuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   26
            Top             =   1530
            Width           =   3570
         End
         Begin TrueOleDBGrid80.TDBGrid grdGrilla 
            Height          =   3105
            Left            =   -74865
            TabIndex        =   44
            Top             =   1365
            Width           =   9840
            _ExtentX        =   17357
            _ExtentY        =   5477
            _LayoutType     =   4
            _RowHeight      =   23
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0).NumberFormat=   "True/False"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   2
            Splits(0).DividerColor=   32768
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
            _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=4,.fontname=Arial"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Arial"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(43)  =   "Named:id=34:Heading"
            _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=34,.fontname=Arial"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
            _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(53)  =   ":id=36,.fontname=Arial"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
            _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid80.TDBGrid grdAI 
            Height          =   3870
            Left            =   -74865
            TabIndex        =   55
            Top             =   1410
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   6826
            _LayoutType     =   4
            _RowHeight      =   23
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0).NumberFormat=   "True/False"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   2
            Splits(0).DividerColor=   32768
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
            _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=4,.fontname=Arial"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Arial"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(43)  =   "Named:id=34:Heading"
            _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=34,.fontname=Arial"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
            _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(53)  =   ":id=36,.fontname=Arial"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
            _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid80.TDBGrid grdAChef 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   173
            Top             =   1620
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5530
            _LayoutType     =   4
            _RowHeight      =   23
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0).NumberFormat=   "True/False"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   2
            Splits(0).DividerColor=   32768
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
            _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=4,.fontname=Arial"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Arial"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(43)  =   "Named:id=34:Heading"
            _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=34,.fontname=Arial"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
            _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(53)  =   ":id=36,.fontname=Arial"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
            _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
         Begin MSComDlg.CommonDialog dlgFoto 
            Left            =   -71880
            Top             =   1860
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog dlgFotoPie 
            Left            =   -68280
            Top             =   1860
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame Frame6 
            Caption         =   " Activa Password "
            Height          =   2685
            Left            =   -69735
            TabIndex        =   89
            Top             =   975
            Width           =   4695
            Begin VB.CheckBox ChkActivaPassworMulticajero 
               Caption         =   "Activa Password MultiCajero Salon, pagos y documentos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   324
               Top             =   2260
               Width           =   4515
            End
            Begin VB.CheckBox chkPassOtrosPagos 
               Caption         =   "Activa Password para Otras Formas de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   243
               Top             =   1960
               Width           =   3585
            End
            Begin VB.CheckBox chkPasswordTransferencia 
               Caption         =   "Activa Password para transferencias"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   95
               Top             =   1143
               Width           =   2985
            End
            Begin VB.CheckBox chkPassword 
               Caption         =   "Activa Password de Eliminacion en el Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   94
               Top             =   591
               Width           =   3780
            End
            Begin VB.CheckBox chkPasswordC 
               Caption         =   "Activa Password de Eliminación de Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   93
               Top             =   315
               Width           =   3690
            End
            Begin VB.CheckBox chkObligaCierre 
               Caption         =   "Activa Password al Cierre del Turno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   92
               Top             =   867
               Width           =   3090
            End
            Begin VB.CheckBox chkPasswordImportar 
               Caption         =   "Activa Password para Importar Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   91
               Top             =   1695
               Width           =   3225
            End
            Begin VB.CheckBox chkPasswordPorCobrar 
               Caption         =   "Activa Password para Enviar a Cuentas Por Cobrar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   90
               Top             =   1419
               Width           =   3945
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   " Activaciones Formas de Venta"
            Height          =   3105
            Left            =   -69735
            TabIndex        =   106
            Top             =   3690
            Width           =   4695
            Begin VB.Frame Frame18 
               Caption         =   "Canales de Venta"
               Height          =   2055
               Left            =   360
               TabIndex        =   275
               Top             =   480
               Visible         =   0   'False
               Width           =   3615
               Begin VB.CheckBox ChCanal 
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   282
                  Top             =   360
                  Width           =   2535
               End
               Begin VB.CheckBox ChCanal 
                  Height          =   255
                  Index           =   2
                  Left            =   360
                  TabIndex        =   281
                  Top             =   600
                  Width           =   2415
               End
               Begin VB.CheckBox ChCanal 
                  Height          =   255
                  Index           =   3
                  Left            =   360
                  TabIndex        =   280
                  Top             =   840
                  Width           =   2655
               End
               Begin VB.CheckBox ChCanal 
                  Height          =   255
                  Index           =   4
                  Left            =   360
                  TabIndex        =   279
                  Top             =   1080
                  Width           =   2535
               End
               Begin VB.CheckBox ChCanal 
                  Height          =   255
                  Index           =   5
                  Left            =   360
                  TabIndex        =   278
                  Top             =   1320
                  Width           =   2655
               End
               Begin VB.CheckBox ChCanalTodos 
                  Caption         =   "Seleccionar Todos"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   277
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Ocultar"
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   276
                  Top             =   1560
                  Width           =   1095
               End
            End
            Begin VB.CheckBox ChkPagoDirectoCD 
               Caption         =   "Activar Pago Directo - Central Delivery"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   120
               TabIndex        =   267
               Top             =   2640
               Width           =   3255
            End
            Begin VB.CommandButton btnVer 
               Caption         =   "Ver Canales de Venta"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3240
               TabIndex        =   266
               Top             =   1680
               Width           =   1335
            End
            Begin VB.CheckBox chkCajaContingencia 
               Caption         =   "Activar Caja Contingencia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   120
               TabIndex        =   239
               Top             =   2400
               Width           =   2520
            End
            Begin VB.CheckBox chkMesa247 
               Caption         =   "Caja Mesa247"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   120
               TabIndex        =   204
               Top             =   2100
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.CheckBox chkWebAp 
               Caption         =   "Apertura integración con app externas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   190
               Top             =   1830
               Width           =   3195
            End
            Begin VB.CheckBox chkCD 
               Caption         =   "Activa Caja Central de Delivery"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   111
               Top             =   250
               Width           =   2730
            End
            Begin VB.CheckBox chkMultiCajero 
               Caption         =   "Activa Multicajero Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   110
               Top             =   800
               Width           =   2610
            End
            Begin VB.CheckBox chkMCPV 
               Caption         =   "Activa Multicajero Salon"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   109
               Top             =   1050
               Width           =   2490
            End
            Begin VB.CheckBox chkCCVOX 
               Caption         =   "Activa Caja Delivery con CCVOX"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   108
               Top             =   510
               Width           =   2730
            End
            Begin VB.CheckBox chkCompatibilidadTVS 
               Caption         =   "Permite Compatibilidad con TVS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   780
               Left            =   3000
               TabIndex        =   107
               Top             =   240
               Visible         =   0   'False
               Width           =   1515
            End
         End
         Begin TrueOleDBGrid80.TDBGrid dgvBloqueoOrigenVenta 
            Height          =   5400
            Left            =   -69120
            TabIndex        =   326
            Top             =   1080
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   9525
            _LayoutType     =   4
            _RowHeight      =   21
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0).NumberFormat=   "True/False"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   2
            Splits(0).DividerColor=   32768
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "Doble click para eliminar"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
            _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=4,.fontname=Arial"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Arial"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(43)  =   "Named:id=34:Heading"
            _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=34,.fontname=Arial"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
            _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(53)  =   ":id=36,.fontname=Arial"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
            _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
         Begin VB.Image imgFotoPie 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   -69480
            Stretch         =   -1  'True
            ToolTipText     =   "Imagen Para Pie de Documentos"
            Top             =   1620
            Width           =   4005
         End
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   -74520
            Stretch         =   -1  'True
            ToolTipText     =   "Imagen Para Cabecera de Documentos"
            Top             =   1620
            Width           =   4005
         End
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo cboPreCuenta 
         Height          =   315
         Left            =   1995
         TabIndex        =   77
         Top             =   555
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1995
         TabIndex        =   78
         Top             =   915
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1995
         TabIndex        =   79
         Top             =   1290
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboSucursal 
         Height          =   315
         Left            =   7695
         TabIndex        =   80
         Top             =   1290
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   88
         Top             =   240
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboUnidadNegocio 
         Height          =   315
         Left            =   7695
         TabIndex        =   141
         Top             =   950
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboSectorVenta 
         Height          =   315
         Left            =   7695
         TabIndex        =   143
         Top             =   600
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboComprobante 
         Height          =   315
         Left            =   7695
         TabIndex        =   218
         Top             =   1650
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   1995
         TabIndex        =   273
         Top             =   1660
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Origen Venta predeterm :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   120
         TabIndex        =   274
         Top             =   1695
         Width           =   1770
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Impresora Comprobante MESA24/7 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   4800
         TabIndex        =   219
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Agrupación de Punto de Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   5265
         TabIndex        =   140
         Top             =   600
         Width           =   2385
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Sucursal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   6120
         TabIndex        =   87
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   6120
         TabIndex        =   86
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Impresora Pre Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   85
         Top             =   675
         Width           =   1620
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   83
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Grupo predeterminado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   82
         Top             =   1005
         Width           =   1755
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido predeterm :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   81
         Top             =   1380
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11895
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9240
      Width           =   11955
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Agregar"
         Height          =   615
         Index           =   0
         Left            =   6960
         Picture         =   "frmCajaDetalle.frx":4D5C
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6660
         TabIndex        =   16
         Top             =   60
         Width           =   6720
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   5220
            Picture         =   "frmCajaDetalle.frx":528E
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   5700
            Picture         =   "frmCajaDetalle.frx":57D0
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   5
            Left            =   6180
            Picture         =   "frmCajaDetalle.frx":5D12
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajaDetalle.frx":6254
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCajaDetalle.frx":6796
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCajaDetalle.frx":6CD8
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            Height          =   195
            Left            =   1590
            TabIndex        =   24
            Top             =   180
            Width           =   3495
         End
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Grabar"
         Height          =   615
         Index           =   1
         Left            =   8140
         Picture         =   "frmCajaDetalle.frx":721A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Eliminar"
         Height          =   615
         Index           =   2
         Left            =   9330
         Picture         =   "frmCajaDetalle.frx":774C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Salir"
         Height          =   615
         Index           =   3
         Left            =   10520
         Picture         =   "frmCajaDetalle.frx":784E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCajaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupo As Recordset
Dim RsGrilla As Recordset
Dim RsImpresora As Recordset
Dim RsPreCuenta As Recordset
Dim RsTipoDocumento As Recordset
Dim RsArea As Recordset
Dim RsTipoPedido As Recordset
Dim RsLocal As Recordset
Dim RsUnidadNegocio As Recordset
Dim rsAreaProduccion As Recordset
Dim RsAI As Recordset
Dim RsImpArea As Recordset
Dim RsFormulario As Recordset
Dim wAgrega As Boolean
Dim Rssucursal As Recordset
Dim RsSectorVenta As Recordset
Dim RsListaAsignada As Recordset
Dim RsOrigenVenta As Recordset
'LG
Dim RsSubGrupo As Recordset
Dim rsAreaProducccionSubgrupo As Recordset
Dim rsAreaSubGrupo As Recordset
Dim strFilenameRuta As String
Dim strFilenameRutaPie As String
'CESAR AREA CHEF
Dim RsAChef As Recordset
Dim RsBloqueoOrigenVenta As New Recordset

Sub LlenaCombos()
    
    With cboTipoDocumento
         Isql = "Select * from vTipoDocumento where lActivo = 1 order by Codigo"
         Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoDocumento
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboFormulario
        If lFESpring Then
            Isql = "Select * from vFormulario where lActivo = 1 and codigo = '01' order by Codigo"
        Else
            Isql = "Select * from vFormulario where lActivo = 1 order by Codigo"
        End If
        Set RsFormulario = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsFormulario
            .DataField = "Descripcion"
            .ListField = "Descripcion"
            .BoundColumn = "Codigo"
    End With
        
    With cboArea
         Isql = "Select * from vArea where lActivo = 1 order by Codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoPedido
         Isql = "Select * from vTipoPedido order by Codigo"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUnidadNegocio
         Isql = "Select * from vUnidadNegocio order by Codigo"
         Set RsUnidadNegocio = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUnidadNegocio
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'multiareas
        With cboAreaProduccion
         Isql = "select * from vArea where lActivo=1 union select 'ABC','SIN AREA','SIN AREA','999',1,'',0,NULL,0,0 order by 1 "
         Set rsAreaProduccion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsAreaProduccion
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'SUCURSALES
        With cboSucursal
         Isql = "Select * from vSucursal where lActivo = 1 order by Codigo"
         Set Rssucursal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = Rssucursal
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboLocal
         Isql = "Select * from vLocal order by Codigo"
         Set RsLocal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsLocal
             .DataField = "Codigo"
             .ListField = "Codigo"
             .BoundColumn = "Codigo"
    End With
    
    'CESAR Sector
    With cboSectorVenta
         Isql = "Select * from vSectorVenta order by Codigo"
         Set RsSectorVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSectorVenta
             .DataField = "Detallado"
             .ListField = "Detallado"
             .BoundColumn = "Codigo"
    End With
    
    With Me.cboSubGrupo
         Isql = "select codigo, tresumido AS descripcion from vSubGrupo where lActivo=1 order by tresumido"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboAreaProd
         Isql = "select * from vArea where lActivo=1 order by 1 "
         Set rsAreaProducccionSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsAreaProducccionSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'CESAR AREA CHEF
    With cboAreaChef
         Isql = "Select * from vArea where lActivo = 1 order by Codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    
End Sub

Sub Asignar()
    txtCodigo = IIf(IsNull(frmCaja.RsCabecera!tCaja), "", frmCaja.RsCabecera!tCaja)
   
    With cboPreCuenta
         Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
         
         Set RsPreCuenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsPreCuenta
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
    
    With cboComprobante
         Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
         
         Set RsPreCuenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsPreCuenta
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
    
    With cboOrigenVenta
         Isql = "Select * from vorigenventa where CodCanalVenta = '" & frmCaja.RsCabecera!tTipoPedido & "' order by CodOrigenVenta"
         
         Set RsOrigenVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigenVenta
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "CodOrigenVenta"
    End With
        
    With frmCaja.RsCabecera
        'Cuadro de Texto
        txtDetallado = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        cboPreCuenta.BoundText = IIf(IsNull(!tPrecuenta), "", Trim(!tPrecuenta))
        cboComprobante.BoundText = IIf(IsNull(!tImpFiscal), "", Trim(!tImpFiscal))
        cboGrupo.BoundText = IIf(IsNull(!tgrupo), "", Trim(!tgrupo))
        cboTipoPedido.BoundText = IIf(IsNull(!tTipoPedido), "", Trim(!tTipoPedido))
        cboUnidadNegocio.BoundText = IIf(IsNull(!tUnidadNegocio), "", Trim(!tUnidadNegocio))
        cboSucursal.BoundText = IIf(IsNull(!tSucursal), "", Trim(!tSucursal))
        cboSectorVenta.BoundText = IIf(IsNull(!tSectorVenta), "", Trim(!tSectorVenta))
        'multiarea
        If IsNull(!tSubAlmacen) Or !tSubAlmacen = "" Then
            cboAreaProduccion.BoundText = "ABC"
        Else
            cboAreaProduccion.BoundText = !tSubAlmacen
        End If
        'Check Box
        chkComanda = IIf(!lComanda = True, 1, 0)
        chkVComanda = IIf(!vComanda = True, 1, 0)
        chkComboPrecuenta = IIf(!lComboPrecuenta = True, 1, 0)
        chkComboDocumento = IIf(!lComboDocumento = True, 1, 0)
        chkAdicion = IIf(!lAdicion = True, 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkConsumo1 = IIf(!lConsumo1 = True, 1, 0)
        chkConsumo2 = IIf(!lConsumo2 = True, 1, 0)
        chkConsumo3 = IIf(!lConsumo3 = True, 1, 0)
        'codigoReciboIngreso
        chkCodigoReciboIngreso = IIf(!lCodigoReciboIngreso = True, 1, 0)
        
        
        chkPrecuenta = IIf(!lPrecuenta = True, 1, 0)
        chkAgrupada = IIf(!lPrecuentaAgrupada = True, 1, 0)
        chkDocumentoAgrupado = IIf(!lDocumentoAgrupado = True, 1, 0)
        chkOrden = IIf(!lOrden = True, 1, 0)
        
        Me.chkBuscaPedido.Value = IIf(!lBuscaPedidoNumero = True, 1, 0)
        chkImprimeImagCabPrecuenta.Value = IIf(IsNull(!lImprimeImagCabPrecuenta), 0, IIf(!lImprimeImagCabPrecuenta = True, 1, 0))
        chkImprimeImagPiePrecuenta.Value = IIf(IsNull(!lImprimeImagPiePrecuenta), 0, IIf(!lImprimeImagPiePrecuenta = True, 1, 0))
        
        chkAccesoDespachoPedido.Value = IIf(IsNull(!lAccesoDespachoPedido), 0, IIf(!lAccesoDespachoPedido = True, 1, 0))
        
        
        chkEliminaC = IIf(!lMotivoEliminaC = True, 1, 0)
        chkPasswordC = IIf(!lPasswordC = True, 1, 0)
        
        chkElimina = IIf(!lMotivoElimina = True, 1, 0)
        chkPassword = IIf(!lPassword = True, 1, 0)
        chkObligaPrinter = IIf(IsNull(!lObligaPrinter), 0, IIf(!lObligaPrinter = True, 1, 0))
        chkObligaPrecuenta = IIf(IsNull(!lObligaPrecuenta), 0, IIf(!lObligaPrecuenta = True, 1, 0))
        chkObligaCierre.Value = IIf(!lObligaCierre = True, 1, 0)
        chkFiltroTipoPedido.Value = IIf(!lFiltroTipoPedido = True, 1, 0)
        chkCancelacion.Value = IIf(!lCancelacion = True, 1, 0)
        chkDirecto.Value = IIf(!lDirecto = True, 1, 0)
        chkCambioMesa.Value = IIf(!lCambioMesa = True, 1, 0)
        txtPuerto = IIf(IsNull(!nPuerto), 0, !nPuerto)
        txtMensaje1 = IIf(IsNull(!tMensaje1), "", !tMensaje1)
        txtMensaje2 = IIf(IsNull(!tMensaje2), "", !tMensaje2)
        txtLongitudBarra = val(IIf(IsNull(!nLongitudBarra), 0, !nLongitudBarra))
        chkVisaNet.Value = IIf(!lVisaNet = True, 1, 0)
        chkImpuestoPrecuenta.Value = IIf(!lImpuestoPrecuenta = True, 1, 0)
        chkValor.Value = IIf(!lValorCortesia = True, 1, 0)
        chkEquivaPrecuenta = IIf(!lEquivaDolaPrecuenta = True, 1, 0)
        chkObservacion.Value = IIf(!lObservacion = True, 1, 0)
        chkCajaRapida.Value = IIf(!lCajaRapida = True, 1, 0)
        chkPropiedadPrecuenta.Value = IIf(!lPropiedadPrecuenta = True, 1, 0)
        chkPropiedadDocumento.Value = IIf(!lPropiedadDocumento = True, 1, 0)
        chkPrecioNetoPrecuenta.Value = IIf(!lPrecioNetoPrecuenta = True, 1, 0)
        
        chkImprimeImagCabPrecuenta.Value = IIf(!lImprimeImagCabPrecuenta = True, 1, 0)
        chkImprimeImagPiePrecuenta.Value = IIf(!lImprimeImagPiePrecuenta = True, 1, 0)
        
        Me.chklBusquedaReserva.Value = IIf(!lBusquedaReservas = True, 1, 0)
        
        Me.chklActivarValesConsumo.Value = IIf(!lActivarValesConsumo = True, 1, 0)
        
        
        'FACTURACION ELECTRONICA
        'chkFacturacionE.value = IIf(!lFacturacionElectronica = True, 1, 0)
        
        txtLimitePrecuenta.Text = IIf(IsNull(!nLimitePrecuenta), 0, !nLimitePrecuenta)
        txtLimiteReimpresion.Text = IIf(IsNull(!nLimiteReimpresion), 0, !nLimiteReimpresion)
        chkPasswordTransferencia.Value = IIf(!lPasswordTransferencia = True, 1, 0)
        chkPasswordImportar.Value = IIf(!lPasswordImportarPedido = True, 1, 0)
        chkCD.Value = IIf(!lCD = True, 1, 0)
        chkMultiCajero.Value = IIf(!lMultiCajero = True, 1, 0)
        chkMCPV.Value = IIf(!lMCPV = True, 1, 0)
        chkCCVOX.Value = IIf(!lCCVOX = True, 1, 0)
        chkObservacionPrecuenta.Value = IIf(!lObservacionPrecuenta = True, 1, 0)
        chkObservacionDocumento.Value = IIf(!lObservacionDocumento = True, 1, 0)
        chkObservacionCabDoc.Value = IIf(!lObservacionCabDoc = True, 1, 0)
        chkDescripcionAlternativa.Value = IIf(!lActivaImpDscAlternativa = True, 1, 0)
        Me.chkMotDesc.Value = IIf(IsNull(!lMotivoDescuento), 0, IIf(!lMotivoDescuento = True, 1, 0))
        Me.chkCajaContingencia.Value = IIf(IsNull(!lCajaContingencia), 0, IIf(!lCajaContingencia = True, 1, 0))
        Me.chkImpPropina.Value = IIf(IsNull(!lImpPropina), 0, IIf(!lImpPropina = True, 1, 0))
        Me.chkComandaF2.Value = IIf(IsNull(!lImpComandaf2), 0, IIf(!lImpComandaf2 = True, 1, 0))
        chkDisgrega.Value = IIf(!lDisgrega = True, 1, 0)
        chkSiab.Value = IIf(!lSiab = True, 1, 0)
        
        chkBloqueaPrecuenta = IIf(!lBloqueaPrecuenta = True, 1, 0)
        If chkBloqueaPrecuenta.Value = 1 Then
                txtLimitePrecuenta.Enabled = False
            Else
                txtLimitePrecuenta.Enabled = True
        End If
        
        'TVS============================
        chkCompatibilidadTVS.Value = IIf(!lCompatibilidadTVS = True, 1, 0)
        '===============================
        
        chkPagoRapido.Value = IIf(!lPagoRapido = True, 1, 0)
        chkPasswordPorCobrar.Value = IIf(!lPasswordPorCobrar = True, 1, 0)
        chkModificaTipoPedido.Value = IIf(!lmodificatipoPedido = True, 1, 0)
        txtBalanzaPuerto = IIf(IsNull(!nBalanzaPuerto), 0, !nBalanzaPuerto)

        chkCajaMobile.Value = IIf(!lCajaMobile = True, 1, 0)
        chkRotulado.Value = IIf(!lRotulado = True, 1, 0)
        
        '0084-2013 CESAR
        chkPagoRapidopv.Value = IIf(!lPagoRapidoPV = True, 1, 0)
        txtTextoConsumo = IIf(IsNull(!tTextoConsumo), "", !tTextoConsumo)
        chkPagoRapidoMod.Value = IIf(!lPagoRapidoMod = True, 1, 0)
        
        chkWebAp.Value = IIf(!lWebAp = True, 1, 0)
        chkMesa247.Value = IIf(!lMesa247 = True, 1, 0)
        
        chkConsumo4.Value = IIf(!lConsumo4 = True, 1, 0)
        
        chkPrecuentaNoValorizada.Value = IIf(!lPrecuentaNoValorizada = True, 1, 0)
        chkEAN13.Value = IIf(!EAN13 = True, 1, 0)
        'ChkBloqueoEdicionPedidoConsolaAgr.Value = IIf(!lBloqueoEdicionPedidoConsolaAgr = True, 1, 0)
         
        Dim xCapturaPeso As Boolean
        xCapturaPeso = IIf(IsNull(!lCapturaPeso), 0, !lCapturaPeso)
        
        If xCapturaPeso = True Then
           opcCapturaPeso.Value = True
           opcCapturaPrecio.Value = False
        Else
           opcCapturaPeso.Value = False
           opcCapturaPrecio.Value = True
        End If
        
        
         If !lMultiAreaSubGrupo = True Then
            chkMulti2.Value = 1
            Me.fra2.Enabled = True
            Else
            chkMulti2.Value = 0
            Me.fra2.Enabled = False
        End If
        
        
        If !lMultiAreaCaja = True Then
            chkMulti1.Value = 1
            Me.fra1.Enabled = True
        Else
            chkMulti1.Value = 0
            Me.fra1.Enabled = False
        End If
        
        'HUELLA
        If !lHuella = True Then
            chkHuella.Value = 1
        Else
            chkHuella.Value = 0
        End If
    
       Me.chkBuscaPedidoVisualizaGrilla.Value = IIf(!lBuscarPedidoVisualizarGrilla = True, 1, 0)
       Me.chkBuscaPedidoFiltrarMesa.Value = IIf(!lBuscarPedidoFiltrarMesa = True, 1, 0)

       Me.chkClaveEnvio.Value = IIf(!lClaveEnvioProduccion = True, 1, 0)
       Me.chkPassOtrosPagos.Value = IIf(!lPassOtrosPagos = True, 1, 0)
       Me.chkImpClienteCab.Value = IIf(!lImpClienteCab = True, 1, 0)
       
       Me.chkBuscarPedidoBarra.Value = IIf(!lBuscarPedidoBarra = True, 1, 0)
       Me.chkIntVisaNetQR.Value = IIf(!lIntVisaNetQR = True, 1, 0)
       
       Me.ChkImpComandaBarra.Value = IIf(!lImpComandaBarra = True, 1, 0)
       Me.ChkBloqEmiDoc.Value = IIf(!lbloqueadocconsumo = True, 1, 0)
       Me.ChkActivaCashDro.Value = IIf(!lCashDro = True, 1, 0)
       Me.chkActivaUber.Value = IIf(!lIntUber = True, 1, 0)
       Me.txtIpCashDrow.Text = IIf(IsNull(!tIpCashDrow), "", Trim(!tIpCashDrow))
       Me.ChkBloqAfectoInafecto.Value = IIf(!lBloqueoAfectoInafecto = True, 1, 0)
       Me.chkIntegracionPatio.Value = IIf(!lIntPatio = True, 1, 0)
       Me.chkIntMercadoPagoQR.Value = IIf(!lIntMercadoPagoQR = True, 1, 0)
       Me.chkIntMercadoPagoQREstatico.Value = IIf(!lIntMercadoPagoQREstatico = True, 1, 0)
       Me.chkPagoEfectivo.Value = IIf(!lIntPagoEfectivo = True, 1, 0)
       Me.ChkPagoDirectoCD.Value = IIf(!lPagoDirectoCD = True, 1, 0)
       Me.chkImprimePropinaSugerida.Value = IIf(!lImprimePropinaSugerida = True, 1, 0)
       Me.chklActivarAvisoRetenciones.Value = IIf(!lActivarAvisoRetenciones = True, 1, 0)
       
       Me.chklActivaTresDecimales.Value = IIf(!lActivaTresDecimales = True, 1, 0)
       Me.chklIntegracionCASSWII30CD.Value = IIf(!lIntegracionCASSWII30CD = True, 1, 0)
       
       If (!tDefaultOrigenVenta <> "") Then
             Me.cboOrigenVenta.BoundText = !tDefaultOrigenVenta
       Else
            Me.cboOrigenVenta.BoundText = ""
       End If
       
       
       Me.ChlNoVisualizarItemCombos.Value = IIf(!lNoVisualizarItemCombos = True, 1, 0)
       Me.ChlMostrarPrecioItemsCombo.Value = IIf(!lMostrarPrecioItemsCombo = True, 1, 0)
       Me.chklActivarMensajeCtsCorrientes.Value = IIf(!lActivarMensajeCtsCorrientes = True, 1, 0)
       Me.chklIntMobileKDS.Value = IIf(!lIntMobileKDS = True, 1, 0)
       Me.chkGuiaF2.Value = IIf(!lImprimirGuiaFormatoCorp = True, 1, 0)
       Me.chklPermitirAnularMaxDias.Value = IIf(!lPermitirAnularMaxDias = True, 1, 0)
       Me.chklOrdenarClientes.Value = IIf(!lOrdenarClientes = True, 1, 0)
        
       
        txtUserToken = IIf(IsNull(!lUserToken), "", !lUserToken)
        txtApiKey = IIf(IsNull(!lApiKey), "", !lApiKey)
        txtApiToken = IIf(IsNull(!lApiToken), "", !lApiToken)

        Me.ChkVistaPreviaItemsCombo.Value = IIf(!lVistaPreviaItemsCombo = True, 1, 0)
        
        Me.chkimprimePropinaDoc.Value = IIf(!lImprimePropinaDoc = True, 1, 0)
        Me.chkIntFPAY.Value = IIf(!IntFPAY = True, 1, 0)
        Me.ChklActivaCreditoCoorporativo.Value = IIf(!lActivaCreditoCoorporativo = True, 1, 0)
        Me.chkIntegracionWebApi.Value = IIf(!IntegracionWebApi = True, 1, 0)
         
        Me.chklModoFacturacion.Value = IIf(!lModoFacturacion = True, 1, 0)
        Me.chklActivarImpresionBlu.Value = IIf(!lActivarImpresionBlu = True, 1, 0)
        Me.ChklActivaRestriClienteDelivery.Value = IIf(!lActivaRestriClienteDelivery = True, 1, 0)
        Me.chkMotivoAdm.Value = IIf(!lMotivoAdm = True, 1, 0)

        Me.ChkBloqueoAnulaPedido.Value = IIf(!lBloqueoAnulaPedido = True, 1, 0)
        Me.ChkBloqueoAnulaItemsPedido.Value = IIf(!lBloqueoAnulaItemsPedido = True, 1, 0)
        
        Me.chkOcultarDesgloseRazonSocial.Value = IIf(!lOcultaRazonSocial = True, 1, 0)
        
        
        Me.ChkBloqueaCanal(1).Value = IIf(!lBloqueaCanal1 = True, 1, 0)
        Me.ChkBloqueaCanal(2).Value = IIf(!lBloqueaCanal2 = True, 1, 0)
        Me.ChkBloqueaCanal(3).Value = IIf(!lBloqueaCanal3 = True, 1, 0)
        Me.ChkBloqueaCanal(4).Value = IIf(!lBloqueaCanal4 = True, 1, 0)
        Me.ChkBloqueaCanal(5).Value = IIf(!lBloqueaCanal5 = True, 1, 0)
        
        Me.chkPrecuentaCabUnica.Value = IIf(!lImprimePrecuentaCabunica = True, 1, 0)
        Me.ChkBloqueaPedidoSinMesa.Value = IIf(!lBloquearPedidoSinMesa = True, 1, 0)
        Me.chkAutorizacionSinMesa.Value = IIf(!lAutorizacionSinMeza = True, 1, 0)
        Me.chkSolicitaClaveEgreso.Value = IIf(!lSolicitaClaveEgreso = True, 1, 0)
        Me.ChkBloqueoEdicionPedidoConsolaAgr.Value = IIf(!lbloqueoEdicionPedidoConsolaAgr = True, 1, 0)
        Me.ChkActivaPassworMulticajero.Value = IIf(!lActivaPasswordMulticajero = True, 1, 0)
        Me.chkPrecuentaImprime(1).Value = IIf(!lImprimePagoPrecuenta = True, 1, 0)
        
        
       On Error GoTo Err
       Dim rst1 As New ADODB.Recordset
       imgFoto.DataField = "foto"
       cmdAgregarFoto.Caption = "Editar Imagen Cabecera"
       Set rst1 = Lib.OpenRecordset("select iimagencabdoc as foto from tcaja where tcaja='" & txtCodigo.Text & "'", Cn)
       Set imgFoto.DataSource = rst1
       imgFotoPie.DataField = "foto"
       cmdAgregarFotoPie.Caption = "Editar Imagen Pie"
       Set rst1 = Lib.OpenRecordset("select iimagenpiedoc as foto from tcaja where tcaja='" & txtCodigo.Text & "'", Cn)
       Set imgFotoPie.DataSource = rst1
       Call asignarBalanza
     
     'Asignacion de compras para uso de datos
     Set dgvBloqueoOrigenVenta.DataSource = Nothing
     Isql = "select o.CanalVenta, CodigoOrigenventa codigo ,o.Descripcion from TCAJAORIGEN_BLOQUEO b inner join vOrigenVenta o on b.CodigoOrigenventa= o.CodOrigenVenta where tCaja='" & txtCodigo.Text & "' order by 1"
     Set RsBloqueoOrigenVenta = Lib.OpenRecordset(Isql, Cn)
     Set dgvBloqueoOrigenVenta.DataSource = RsBloqueoOrigenVenta
     
Err:
       
    End With
    
    'Cambiar el Filtro
    RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
    RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
    'CESAR AREA CHEF
    RsAChef.Filter = "tCaja ='" & txtCodigo.Text & "'"
    '----------------------
    rsAreaSubGrupo.Filter = "tCaja ='" & txtCodigo.Text & "'"
End Sub

Private Sub btnNewModal_Click()
    If txtCodigo <> "" Then
        frmCajaDetalle2.CodigoCaja = txtCodigo
        frmCajaDetalle2.DesCaja = txtDetallado
        frmCajaDetalle2.Show vbModal
    Else
        MsgBox "No tiene una caja asignada, recuerda que primero debes registrar en mantenimiento de cajas para realizar configuración adicionales.", vbInformation, sMensaje
    End If
End Sub

Private Sub BtnVer_Click()
    Frame18.Visible = True
    
      Isql = " Select   tcaja,   tcanalventa1,  tcanalventa2,  tcanalventa3,  tcanalventa4,  tcanalventa5," & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '01' )acti1, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '02' )acti2, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '03' )acti3, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '04' )acti4, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '05' )acti5 " & _
           "  from TCAJACANALVENTA where tCaja = '" & txtCodigo.Text & "'"
           Set RsListaAsignada = Lib.OpenRecordset(Isql, Cn)
    
    If Not (RsListaAsignada.EOF Or RsListaAsignada.BOF) Then
        RsListaAsignada.MoveFirst
        With RsListaAsignada
        
             Me.ChCanal(1).Caption = IIf(IsNull(!acti1), "", !acti1)
             Me.ChCanal(2).Caption = IIf(IsNull(!acti2), "", !acti2)
             Me.ChCanal(3).Caption = IIf(IsNull(!acti3), "", !acti3)
             Me.ChCanal(4).Caption = IIf(IsNull(!acti4), "", !acti4)
             Me.ChCanal(5).Caption = IIf(IsNull(!acti5), "", !acti5)
                                 
             Me.ChCanal(1).Value = IIf(IsNull(!tcanalventa1), 0, IIf(!tcanalventa1 = True, 1, 0))
             Me.ChCanal(2).Value = IIf(IsNull(!tcanalventa2), 0, IIf(!tcanalventa2 = True, 1, 0))
             Me.ChCanal(3).Value = IIf(IsNull(!tcanalventa3), 0, IIf(!tcanalventa3 = True, 1, 0))
             Me.ChCanal(4).Value = IIf(IsNull(!tcanalventa4), 0, IIf(!tcanalventa4 = True, 1, 0))
             Me.ChCanal(5).Value = IIf(IsNull(!tcanalventa5), 0, IIf(!tcanalventa5 = True, 1, 0))
        End With
    Else
             Me.ChCanal(1).Caption = ""
             Me.ChCanal(2).Caption = ""
             Me.ChCanal(3).Caption = ""
             Me.ChCanal(4).Caption = ""
             Me.ChCanal(5).Caption = ""

             Me.ChCanal(1).Value = False
             Me.ChCanal(2).Value = False
             Me.ChCanal(3).Value = False
             Me.ChCanal(4).Value = False
             Me.ChCanal(5).Value = False
    End If
    
    
    
    
     If ChCanal(1).Caption = "" Then
            Me.ChCanal(1).Enabled = False
        End If
        If ChCanal(2).Caption = "" Then
             Me.ChCanal(2).Enabled = False
        End If
        If ChCanal(4).Caption = "" Then
            Me.ChCanal(3).Enabled = False
        End If
        If ChCanal(4).Caption = "" Then
            Me.ChCanal(4).Enabled = False
        End If
        If ChCanal(5).Caption = "" Then
            Me.ChCanal(5).Enabled = False
        End If
    
End Sub

Private Sub cboTipoDocumento_Change()
    If cboTipoDocumento.BoundText = "00" Then
       'Label(4).Visible = False
       Frame14.Visible = False
    Else
       'Label(4).Visible = True
       Frame14.Visible = True
    End If
    
    If sImpuesto1 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto1.Visible = True
       chkImpuesto1.Caption = sImpuesto1
       chkImpuesto1.Value = 1
    Else
       chkImpuesto1.Visible = False
       chkImpuesto1.Value = 0
    End If
        
    If sImpuesto2 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto2.Visible = True
       chkImpuesto2.Caption = sImpuesto2
       chkImpuesto2.Value = 1
    Else
       chkImpuesto2.Visible = False
       chkImpuesto2.Value = 0
    End If
        
    If sImpuesto3 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto3.Visible = True
       chkImpuesto3.Caption = sImpuesto3
       chkImpuesto3.Value = 1
    Else
      chkImpuesto3.Visible = False
      chkImpuesto3.Value = 0
    End If
End Sub

Private Sub cboTipoPedido_Change()
On Error GoTo fin:

    With cboOrigenVenta
         Isql = "Select * from vorigenventa where CodCanalVenta = '" & Me.cboTipoPedido.BoundText & "' order by CodOrigenVenta"
         Set RsOrigenVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigenVenta
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "CodOrigenVenta"
    End With

    Me.cboOrigenVenta.BoundText = ""

Exit Sub
fin:
End Sub

Private Sub ChCanalTodos_Click()

    If ChCanalTodos.Value = 1 Then
        If ChCanal(1).Caption <> "" Then
            Me.ChCanal(1).Value = 1
        End If
        If ChCanal(2).Caption <> "" Then
             Me.ChCanal(2).Value = 1
        End If
        If ChCanal(3).Caption <> "" Then
            Me.ChCanal(3).Value = 1
        End If
        If ChCanal(4).Caption <> "" Then
            Me.ChCanal(4).Value = 1
        End If
        If ChCanal(5).Caption <> "" Then
            Me.ChCanal(5).Value = 1
        End If
     
    Else
            Me.ChCanal(1).Value = 0
             Me.ChCanal(2).Value = 0
             Me.ChCanal(3).Value = 0
             Me.ChCanal(4).Value = 0
             Me.ChCanal(5).Value = 0
    End If
    
       
End Sub



Private Sub chkAgrupada_Click()
   If chkAgrupada.Value Then
      chkComboPrecuenta.Value = False
      chkComboPrecuenta.Enabled = False
      chkPropiedadPrecuenta.Enabled = False
      chkPropiedadPrecuenta.Value = 0
      chkObservacionPrecuenta.Value = 0
      chkObservacionPrecuenta.Enabled = False
      chkPrecuentaCabUnica.Enabled = True
      
      
   Else
      chkPropiedadPrecuenta.Enabled = True
      chkComboPrecuenta.Enabled = True
      chkObservacionPrecuenta.Enabled = True
      chkPrecuentaCabUnica.Enabled = False
      chkPropiedadPrecuenta.Value = 0
      chkObservacionPrecuenta.Value = 0
      
      If chkImpuestoPrecuenta.Value = 0 And chkEquivaPrecuenta.Value = 0 And chkPrecuentaNoValorizada.Value = 0 Then
        chkPrecuentaCabUnica.Enabled = False
        chkPrecuentaCabUnica.Value = False
      Else
        chkPrecuentaCabUnica.Enabled = True
      End If
   End If
End Sub

Private Sub chkBloqueaPrecuenta_Click()
    If chkBloqueaPrecuenta.Value = 1 Then
            If chkObligaPrecuenta.Value = 1 Then
                MsgBox "Está activada la Obligatoriedad de Emisión de Precuentas", vbCritical, sMensaje
                chkBloqueaPrecuenta.Value = 0
                Exit Sub
            Else
        
                txtLimitePrecuenta.Enabled = False
            End If
        Else
            txtLimitePrecuenta.Enabled = True
    End If
End Sub




Private Sub chkDocumentoAgrupado_Click()
   If chkDocumentoAgrupado.Value Then
      chkComboDocumento.Value = False
      chkComboDocumento.Enabled = False
      
      chkPropiedadDocumento.Enabled = False
      chkObservacionDocumento.Enabled = False
      ChlMostrarPrecioItemsCombo.Enabled = False
      chkPropiedadDocumento.Value = 0
      chkObservacionDocumento.Value = 0
      ChlMostrarPrecioItemsCombo.Value = 0
   Else
      chkComboDocumento.Enabled = True
      chkPropiedadDocumento.Enabled = True
      chkObservacionDocumento.Enabled = True
      ChlMostrarPrecioItemsCombo.Enabled = True
      chkPropiedadDocumento.Value = 0
      chkObservacionDocumento.Value = 0
      ChlMostrarPrecioItemsCombo.Value = 0
   End If
End Sub

Private Sub chkEquivaPrecuenta_Click()
    If chkEquivaPrecuenta.Value Then
        chkPrecuentaCabUnica.Enabled = True
    Else
      If chkAgrupada.Value = 0 And chkImpuestoPrecuenta.Value = 0 And chkPrecuentaNoValorizada.Value = 0 Then
        chkPrecuentaCabUnica.Enabled = False
        chkPrecuentaCabUnica.Value = False
    Else
        chkPrecuentaCabUnica.Enabled = True
      End If
   End If
End Sub

Private Sub chkImpuestoPrecuenta_Click()
    If chkImpuestoPrecuenta.Value Then
        chkPrecuentaCabUnica.Enabled = True
    Else
      If chkAgrupada.Value = 0 And chkEquivaPrecuenta.Value = 0 And chkPrecuentaNoValorizada.Value = 0 Then
        chkPrecuentaCabUnica.Enabled = False
        chkPrecuentaCabUnica.Value = False
    Else
        chkPrecuentaCabUnica.Enabled = True
      End If
   End If
End Sub

Private Sub chkMCPV_Click()
  If chkMCPV.Value = 1 Then
     chkMultiCajero.Enabled = False
     chkMultiCajero.Value = 0
  Else
     chkMultiCajero.Enabled = True
  End If
End Sub

Private Sub chkMesa247_Click()
    If chkMesa247.Value = 1 Then
        Label(25).Visible = True
        cboComprobante.Visible = True
    Else
        Label(25).Visible = False
        cboComprobante.Visible = False
    End If
End Sub

Private Sub chkMultiCajero_Click()
  If chkMultiCajero.Value = 1 Then
     chkMCPV.Enabled = False
     chkMCPV.Value = 0
  Else
     chkMCPV.Enabled = True
  End If
End Sub

Private Sub chkObligaPrecuenta_Click()
    If chkObligaPrecuenta.Value = 1 Then
        If chkBloqueaPrecuenta.Value = 1 Then
            MsgBox "La Emisión de Precuentas esta bloqueada", vbCritical, sMensaje
            chkObligaPrecuenta.Value = 0
            Exit Sub
        End If
    End If
        
End Sub

Private Sub chkPrecioNetoPrecuenta_Click()
    If chkPrecioNetoPrecuenta.Value = 1 Then
        chkImpuestoPrecuenta.Value = 1
        chkImpuestoPrecuenta.Enabled = False
    Else
        chkImpuestoPrecuenta.Enabled = True
    End If
End Sub

Private Sub chkPrecuentaNoValorizada_Click()
    If chkPrecuentaNoValorizada.Value Then
        chkPrecuentaCabUnica.Enabled = True
    Else
      If chkAgrupada.Value = 0 And chkImpuestoPrecuenta.Value = 0 And chkEquivaPrecuenta.Value = 0 Then
        chkPrecuentaCabUnica.Enabled = False
        chkPrecuentaCabUnica.Value = False
    Else
        chkPrecuentaCabUnica.Enabled = True
      End If
   End If
End Sub

Private Sub chkWebAp_Click()
' GCAA 26102020

If chkWebAp.Value = 1 Then
    BtnVer.Enabled = True
Else
     BtnVer.Enabled = False
End If





 Dim RsBalanza As Recordset
    Isql = "Select  count(*) cantidad from TCAJACANALVENTA where tCaja = '" & txtCodigo.Text & "'"
    Set RsBalanza = Lib.OpenRecordset(Isql, Cn)
   ' MsgBox (RsBalanza!Cantidad)

    If RsBalanza!Cantidad = 0 Then
         Isql = "insert into TCAJACANALVENTA( " & _
                          "tCaja, tCanalventa1, tCanalventa2, tCanalventa3, tCanalventa4, tCanalventa5, tUsuario, fRegistro ) " & _
                          "values (  '" & txtCodigo.Text & "', 0,0,0,0,0 , " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
        Cn.Execute Isql
    End If
    
    Isql = " Select   tcaja,   tcanalventa1,  tcanalventa2,  tcanalventa3,  tcanalventa4,  tcanalventa5," & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '01' )acti1, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '02' )acti2, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '03' )acti3, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '04' )acti4, " & _
           " ( select tDetallado from tcanalventa  where tCodigoCanalVenta = '05' )acti5 " & _
           "  from TCAJACANALVENTA where tCaja = '" & txtCodigo.Text & "'"
           Set RsListaAsignada = Lib.OpenRecordset(Isql, Cn)
    
    If Not (RsListaAsignada.EOF Or RsListaAsignada.BOF) Then
        RsListaAsignada.MoveFirst
        With RsListaAsignada
        
             Me.ChCanal(1).Caption = IIf(IsNull(!acti1), "", !acti1)
             Me.ChCanal(2).Caption = IIf(IsNull(!acti2), "", !acti2)
             Me.ChCanal(3).Caption = IIf(IsNull(!acti3), "", !acti3)
             Me.ChCanal(4).Caption = IIf(IsNull(!acti4), "", !acti4)
             Me.ChCanal(5).Caption = IIf(IsNull(!acti5), "", !acti5)
                                 
             Me.ChCanal(1).Value = IIf(IsNull(!tcanalventa1), 0, IIf(!tcanalventa1 = True, 1, 0))
             Me.ChCanal(2).Value = IIf(IsNull(!tcanalventa2), 0, IIf(!tcanalventa2 = True, 1, 0))
             Me.ChCanal(3).Value = IIf(IsNull(!tcanalventa3), 0, IIf(!tcanalventa3 = True, 1, 0))
             Me.ChCanal(4).Value = IIf(IsNull(!tcanalventa4), 0, IIf(!tcanalventa4 = True, 1, 0))
             Me.ChCanal(5).Value = IIf(IsNull(!tcanalventa5), 0, IIf(!tcanalventa5 = True, 1, 0))
        End With
    Else
             Me.ChCanal(1).Caption = ""
             Me.ChCanal(2).Caption = ""
             Me.ChCanal(3).Caption = ""
             Me.ChCanal(4).Caption = ""
             Me.ChCanal(5).Caption = ""

             Me.ChCanal(1).Value = False
             Me.ChCanal(2).Value = False
             Me.ChCanal(3).Value = False
             Me.ChCanal(4).Value = False
             Me.ChCanal(5).Value = False
    End If
    
    
    
    

End Sub

Private Sub cmdAgregarFoto_Click()
On Error GoTo errHandler

If txtCodigo.Text <> "" Then
    dlgFoto.CancelError = False
    With cmdAgregarFoto
        If .Caption = "Editar Imagen Cabecera" Then
            dlgFoto.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFoto.FileName = ""
            dlgFoto.ShowOpen
            imgFoto.Visible = True
            If dlgFoto.FileName <> "" Then
                .Caption = "Guardar Imagen Cabecera"
                strFilenameRuta = dlgFoto.FileName
                imgFoto.Picture = LoadPicture(strFilenameRuta)
            End If
        Else
            
             GuardarFoto strFilenameRuta, "1"
            .Caption = "Editar Imagen Cabecera"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para la Caja"
    Exit Sub
End If
errHandler:
MsgBox "Dimensiones Recomendada para la Imagen 350*250 pixeles"
strFilenameRuta = ""
cmdAgregarFoto.Caption = "Editar Imagen Cabecera"
    Exit Sub
End Sub


Public Sub GuardarFoto(ByVal Ruta As String, ByVal tTipo As String)
        
        Dim imgTeacher()      As Byte
        Dim varPhoto          As Variant
        Dim numfile           As Long
        If (Ruta <> "") Then
            varPhoto = FileLen(Ruta)
            ReDim bufimages(varPhoto - 1) As Byte
            numfile = FreeFile
            Open Ruta For Binary As #numfile
            Get #numfile, , bufimages
            Close #numfile
             imgTeacher = bufimages
        End If
        If (Ruta = "") Then
            imgTeacher = LoadResData(101, "CUSTOM")
            varPhoto = UBound(imgTeacher)
        End If
        Dim lnfoto As Variant
        lnfoto = varPhoto
        Dim Cmd As New ADODB.Command
        Dim prm As New ADODB.Parameter
        With Cmd
                .ActiveConnection = Cn
                .CommandText = "sp_UpdImagenCaja"
                .CommandType = adCmdStoredProc
        End With
        Set prm = Cmd.CreateParameter("@tCodigo", adChar, adParamInput, 10, txtCodigo.Text)
        Cmd.Parameters.Append prm
         Set prm = Cmd.CreateParameter("@tTipo", adChar, adParamInput, 1, tTipo)
        Cmd.Parameters.Append prm
        Set prm = Cmd.CreateParameter("@oFoto", adLongVarBinary, adParamInput, lnfoto + 1)
        Cmd.Parameters.Append prm
        
        If Not IsNull(imgTeacher) Then
            prm.AppendChunk imgTeacher
        Else
            prm.Value = Null
        End If
        Cmd.Execute
End Sub

Private Sub cmdAgregarFotoPie_Click()
'on error GoTo ErrHandler
If txtCodigo.Text <> "" Then
    dlgFotoPie.CancelError = False
    With cmdAgregarFotoPie
        If .Caption = "Editar Imagen Pie" Then
            dlgFotoPie.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFotoPie.FileName = ""
            dlgFotoPie.ShowOpen
            imgFotoPie.Visible = True
            If dlgFotoPie.FileName <> "" Then
                .Caption = "Guardar Imagen Pie"
                strFilenameRutaPie = dlgFotoPie.FileName
                imgFotoPie.Picture = LoadPicture(strFilenameRutaPie)
            End If
        Else
            
             GuardarFoto strFilenameRutaPie, "2"
            .Caption = "Editar Imagen Pie"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para la Caja"
    Exit Sub
End If
errHandler:
MsgBox "Dimensiones Recomendada para la Imagen 350*250 pixeles"
strFilenameRutaPie = ""
cmdAgregarFoto.Caption = "Editar Imagen Pie"
End Sub

Private Sub cmdGuardarBalanza_Click()
On Error GoTo fin
    If Calcular("select COUNT(*) AS CODIGO from TCONFIGURAPERIFERICO where tcaja='" & txtCodigo.Text & "' and tTabla='BALANZA'", Cn) = 0 Then
        Isql = "insert into TCONFIGURAPERIFERICO (tTabla,tcaja,nDato1,nDato2,nDato3,nDato4,nDato5,nDato6,nDato7,nDato8, lActivo)" & _
                "Values ('BALANZA', '" & txtCodigo.Text & "', '" & Me.cboBal1.Text & "', '" & Me.cboBal2.Text & "', '" & Me.cboBal3.Text & _
                "', '" & Me.cboBal4.Text & "', '" & Me.cboBal5.Text & "', '" & txtBalanzaPuerto.Text & "', '" & Me.txtBalcomando.Text & "', '" & Me.txtbaltiempo.Text & "'," & Me.chkActivoBal.Value & ")"
                
        Cn.Execute Isql
    Else
        Isql = "update TCONFIGURAPERIFERICO set nDato1='" & Me.cboBal1.Text & "',nDato2='" & Me.cboBal2.Text & "',nDato3='" & Me.cboBal3.Text & _
        "',nDato4='" & Me.cboBal4.Text & "',nDato5='" & Me.cboBal5.Text & "',nDato6='" & txtBalanzaPuerto.Text & "',nDato7='" & Me.txtBalcomando.Text & "'" & _
        ",nDato8='" & Me.txtbaltiempo.Text & "', lActivo= " & Me.chkActivoBal.Value & "  where tTabla='BALANZA' and tcaja='" & txtCodigo.Text & "'"
        Cn.Execute Isql
    End If
    MsgBox "Proceso Correcto", vbInformation, "Inforest"
Exit Sub
fin:
End Sub
Private Sub asignarBalanza()
On Error GoTo fin
    Dim RsBalanza As Recordset
    Isql = "Select * from vBalanza where tCaja = '" & txtCodigo.Text & "'"
    Set RsBalanza = Lib.OpenRecordset(Isql, Cn)
    
    If Not (RsBalanza.EOF Or RsBalanza.BOF) Then
        RsBalanza.MoveFirst
        With RsBalanza
             Me.cboBal1.Text = IIf(IsNull(!nDato1), "", !nDato1)
             Me.cboBal2.Text = IIf(IsNull(!nDato2), "", !nDato2)
             Me.cboBal3.Text = IIf(IsNull(!nDato3), "", !nDato3)
             Me.cboBal4.Text = IIf(IsNull(!nDato4), "", !nDato4)
             Me.cboBal5.Text = IIf(IsNull(!nDato5), "", !nDato5)
             Me.txtBalanzaPuerto.Text = IIf(IsNull(!nDato6), "", !nDato6)
             Me.txtBalcomando.Text = IIf(IsNull(!nDato7), "", !nDato7)
             Me.txtbaltiempo.Text = IIf(IsNull(!nDato8), "", !nDato8)
             Me.chkActivoBal.Value = IIf(IsNull(!lActivo), 0, IIf(!lActivo = True, 1, 0))
        End With
    Else
         Me.cboBal1.Text = ""
         Me.cboBal2.Text = ""
         Me.cboBal3.Text = ""
         Me.cboBal4.Text = ""
         Me.cboBal5.Text = ""
         Me.txtBalanzaPuerto.Text = ""
         Me.txtbaltiempo.Text = ""
         Me.txtBalcomando.Text = ""
         Me.chkActivoBal.Value = 0
    End If
Exit Sub
fin:

End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCaja.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCaja.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCaja.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCaja.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCaja.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCaja.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               inicio
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                                   
               If Sw Then
                  Sw = False
                  
                  'Asignar El Campo de Codificación
                  nCorrela = Calcular("select max(tCaja) as Codigo from TCAJA", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigo = "001"
                  Else
                     txtCodigo = Lib.Correlativo(nCorrela, 3)
                  End If
                                     
               sPasa = txtCodigo.Text
               
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "01", sUsuario, sPasa, "", _
                            "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text, "tPrecuenta", "Impresora Precuenta", cboPreCuenta.BoundText, "lSiab", "Flag Activacion enlace Sistema SIAB", IIf(chkSiab.Value, "Verdadero", "Falso"), _
                            "tGrupo", "Grupo Predeterminado", cboGrupo.BoundText, "tTipoPedido", "Tipo Pedido Predeterminado", cboTipoPedido.BoundText, _
                            "tUnidadNegocio", "Unidad de Negocio", cboUnidadNegocio.BoundText, "tSucursal", "Sucursal", cboSucursal.BoundText, _
                            "tSubAlmacen", "Area de Produccion", IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), _
                            "nPuerto", "Visor Puerto", val(txtPuerto.Text), "tMensaje1", "Visor Mensaje 1", txtMensaje1.Text, "tMensaje2", "Visor Mensaje 2", txtMensaje2.Text, _
                            "nBalanzaPuerto", "Balanza Electronica Puerto", val(txtBalanzaPuerto.Text), "lVisaNet", "Flag Enlace VisaNet", IIf(chkVisaNet.Value, "Verdadero", "Falso"), "nLongitudBarra", "Lector Barra Longitud", txtLongitudBarra.Text, "lcapturaPeso", "Lector de Barra", IIf(opcCapturaPeso.Value, "Verdadero", "Falso"), _
                            "lComanda", "Flag Obligatoriedad Comanda Manual", IIf(chkComanda.Value, "Verdadero", "Falso"), "lObligaPrinter", "Flag Obliga Impresion Pedido", IIf(chkObligaPrinter.Value, "Verdadero", "Falso"), "lObligaPrecuenta", "Flag Obliga Impresion Precuenta", IIf(chkObligaPrecuenta.Value, "Verdadero", "Falso"), "lCancelacion", "Flag Obligatoriedad de Cancelacion", IIf(chkCancelacion.Value, "Verdadero", "Falso"), "lObservacion", "Flag Obligatoriedad Observacion", IIf(chkObservacion.Value, "Verdadero", "Falso"), "lConsumo1", "Flag Emision Rapida por Consumo", IIf(chkConsumo1.Value, "Verdadero", "Falso"), "lConsumo2", "Flag Emision por Consumo En Pagos y Division", IIf(chkConsumo2.Value, "Verdadero", "Falso"), "lConsumo3", "Flag Emision por Consumo en Caja Rapida", IIf(chkConsumo3.Value, "Verdadero", "Falso"), _
                            "lPrecuentaAgrupada", "Flag Impresion Agrupada en Precuenta", IIf(chkAgrupada.Value, "Verdadero", "Falso"), "lDocumentoAgrupado", "Flag Impresion Agrupada en Documentos", IIf(chkDocumentoAgrupado.Value, "Verdadero", "Falso"), "lComboPrecuenta", "Flag Impresion de Combos en Precuenta", IIf(chkComboPrecuenta.Value, "Verdadero", "Falso"), "lComboDocumento", "Flag Impresion de Combos en Documentos", IIf(chkComboDocumento.Value, "Verdadero", "Falso"), "lPropiedadPrecuenta", "Flag Impresion Propiedad en Precuenta", IIf(chkPropiedadPrecuenta.Value, "Verdadero", "Falso"), "lObservacionPrecuenta", "Flag Impresion de Observacion en Precuenta", IIf(chkObservacionPrecuenta.Value, "Verdadero", "Falso"), "lPropiedadDocumento", "Flag Impresion Propiedad en Documento", IIf(chkPropiedadDocumento.Value, "Verdadero", "Falso"), "lObservacionDocumento", "Flag Impresion Observacion en Documento", IIf(chkObservacionDocumento.Value, "Verdadero", "Falso"), _
                            "lPrecioNetoPrecuenta", "Flag Impresion Prec Neto en Precuenta", IIf(chkPrecioNetoPrecuenta.Value, "Verdadero", "Falso"), "lPrecuenta", "Flag Permite Cambiar Impresora Precuenta", IIf(chkPrecuenta.Value, "Verdadero", "Falso"), "lImpuestoPrecuenta", "Flag Impresion Impuesto Desglos. Precuenta", IIf(chkImpuestoPrecuenta.Value, "Verdadero", "Falso"), "lCambioMesa", "Flag Impresion de Cambio de Mesa", IIf(chkCambioMesa.Value, "Verdadero", "Falso"), "lValorCortesia", "Flag Impresion Valorizada de Cortesias", IIf(chkValor.Value, "Verdadero", "Falso"), "lequivadolaprecuenta", "Flag Impresion de Equivalencia Dolares en Precuenta", IIf(chkEquivaPrecuenta.Value, "Verdadero", "Falso"), "lActivaImpDscAlternativa", "Flag Impresion Descripcion Alternativa", IIf(chkDescripcionAlternativa.Value, "Verdadero", "Falso"), "nLimitePrecuenta", "Limite de Precuentas", val(txtLimitePrecuenta.Text), "nLimiteReimpresion", "Limite de Re Impresiones Pedido", val(txtLimiteReimpresion.Text), _
                            "vComanda", "Flag Activa Ingreso Comanda Manual", IIf(chkVComanda.Value, "Verdadero", "Falso"), "lMotivoEliminaC", "Flag Pide Motivo Elimina Pedido", IIf(chkEliminaC.Value, "Verdadero", "Falso"), "lPasswordC", "Flag Activa Password Eliminacion Pedido", IIf(chkPasswordC.Value, "Verdadero", "Falso"), "lMotivoElimina", "Flag Pide Motivo Elimina Producto", IIf(chkElimina.Value, "Verdadero", "Falso"), "lPassword", "Flag Activa Password Elimina Producto", IIf(chkPassword.Value, "Verdadero", "Falso"), "lObligacierre", "Flag Activa Password Cierre Turno", IIf(chkObligaCierre.Value, "Verdadero", "Falso"), "lPasswordTransferencia", "Flag Activa Password Transferencia", IIf(chkPasswordTransferencia.Value, "Verdadero", "Falso"), "lpasswordporcobrar", "Flag Activa Password Por Cobrar", IIf(chkPasswordPorCobrar.Value, "Verdadero", "Falso"), "lPasswordImportarPedido", "Flag Activa Password Importar Pedido", IIf(chkPasswordImportar.Value, "Verdadero", "Falso"), _
                            "lFiltroTipoPedido", "Flag Permite Importar Pedidos por Canal", IIf(chkFiltroTipoPedido.Value, "Verdadero", "Falso"), "lAdicion", "Flag Permite Transferencias", IIf(chkAdicion.Value, "Verdadero", "Falso"), "lmodificatipopedido", "Flag Permite Modificar Tipo de Pedido", IIf(chkModificaTipoPedido.Value, "Verdadero", "Falso"), "lCajaRapida", "Flag Ingreso Directo a Caja Rapida", IIf(chkCajaRapida.Value, "Verdadero", "Falso"), "lPagoRapido", "Flag Ingreso a Pago Rapido desde Caja Rapida", IIf(chkPagoRapido.Value, "Verdadero", "Falso"), "lOrden", "Flag Activa Control Enum. Automatica", IIf(chkOrden.Value, "Verdadero", "Falso"), "lDirecto", "Flag Activa Control de Envios Directos", IIf(chkDirecto.Value, "Verdadero", "Falso"), "lDisgrega", "Flag Disgregar en Dos Partes", IIf(chkDisgrega.Value, "Verdadero", "Falso"), "lCD", "Flag Activa Caja Central Delivery", IIf(chkCD.Value, "Verdadero", "Falso"), "lCCVOX", "Flag Activa Caja Delivery CCVOX", IIf(chkCCVOX.Value, "Verdadero", "Falso"), _
                            "lMultiCajero", "Flag Activa Multicajero Caja Rapida", IIf(chkMultiCajero.Value, "Verdadero", "Falso"), "lMCPV", "Flag Activa Multicajero Salon", IIf(chkMCPV.Value, "Verdadero", "Falso"), "lCompatibilidadTVS", "Flag Permite Compatibilidad con TVS", IIf(chkCompatibilidadTVS.Value, "Verdadero", "Falso"), "lPagoRapidoPV", "Flag Ingreso a Pago Rapido desde Punto Venta", IIf(chkPagoRapidopv.Value, "Verdadero", "Falso"), "tTextoConsumo", "Motivo de Consumo Predeterminado", txtTextoConsumo.Text, "tSectorVenta", "SectorVenta", cboSectorVenta.BoundText, "lCajaMobile", "Flag Caja Mobile", IIf(chkCajaMobile.Value, "Verdadero", "Falso"), "lBloqueaPrecuenta", "Bloquear Precuenta", IIf(chkBloqueaPrecuenta.Value, "Verdadero", "Falso"), "lRotulado", "Enlace Rotulado", IIf(chkRotulado.Value, "Verdadero", "Falso"), "lMultiAreaSubGrupo", "Flag Multi Area Por SubGrupo", IIf(Me.chkMulti2.Value, "Verdadero", "Falso"), _
                            "lMultiAreaCaja", "Flag Multi Area ", IIf(Me.chkMulti1.Value, "Verdadero", "Falso"), "lHuella", "Flag Huella ", IIf(Me.chkHuella.Value, "Verdadero", "Falso"), "lImprimeImagCabPrecuenta", "Imagen Cabecera Precuenta", IIf(Me.chkImprimeImagCabPrecuenta.Value, "Verdadero", "Falso"), "lImprimeImagpiePrecuenta", "Imagen Pie Precuenta", IIf(Me.chkImprimeImagPiePrecuenta.Value, "Verdadero", "Falso"), "lAccesoDespachoPedido", "Acceso Despacho Pedido", IIf(Me.chkAccesoDespachoPedido.Value, "Verdadero", "Falso"), "LBuscarpedidovisualizargrilla", "Buscar Pedido Visualizar Grilla", IIf(Me.chkBuscaPedidoVisualizaGrilla.Value, "Verdadero", "Falso"), "lbuscarpedidofiltrarmesa", "Buscar Pedido Filtrar Mesa", IIf(Me.chkBuscaPedidoFiltrarMesa.Value, "Verdadero", "Falso"), "lMotivoDescuento", "Imprime Motivo Descuento", IIf(Me.chkMotDesc.Value, "Verdadero", "Falso"), "lCajaContingencia", "Activa Caja Contingencia", IIf(Me.chkCajaContingencia.Value, "Verdadero", "Falso"), _
                            "lImpPropina", "Solicita Propina Imp Prec", IIf(Me.chkImpPropina.Value, "Verdadero", "Falso"), "lImpComandaf2", "imprime comanda formato 2", IIf(Me.chkComandaF2.Value, "Verdadero", "Falso"), "lPassOtrosPagos", "Activa Password otros pagos", IIf(Me.chkPassOtrosPagos.Value, "Verdadero", "Falso"), "lImpClienteCab", "Imprime Cliente en Cabecera", IIf(Me.chkImpClienteCab.Value, "Verdadero", "Falso"), "lBuscarPedidoBarra", "Buscar pedido no enviados por codigo de barra", IIf(chkBuscarPedidoBarra.Value, "Verdadero", "Falso"), "lIntVisaNetQR", "IntegracionVisaNet QR", IIf(Me.chkIntVisaNetQR.Value, "Verdadero", "Falso"), "lImpComandaBarra", "Imprime comanda con barra por Item", IIf(Me.ChkImpComandaBarra.Value, "Verdadero", "Falso"), "lIntUber", "Activa Integracion Uber", IIf(Me.chkActivaUber.Value, "Verdadero", "Falso"), "tIpCashDrow", "Ip para la conexion a Cash Drow", Me.txtIpCashDrow.Text, "lBloqueoAfectoInafecto", "Activa el proceso de bloqueo de afectos e inafectos en documentos", _
                            IIf(Me.ChkBloqAfectoInafecto.Value, "Verdadero", "Falso"), "lIntMercadoPagoQR", "Activa Mercado Pago Estatico", IIf(Me.chkIntMercadoPagoQREstatico.Value, "Verdadero", "Falso"), "lIntPagoEfectivo", "INtegracion pago Efectivo", IIf(Me.chkPagoEfectivo.Value, "Verdadero", "Falso"), "lPagoDirectoCD", "Pago Directo Central Delivery", IIf(Me.ChkPagoDirectoCD.Value, "Verdadero", "Falso"), "lImprimePropinaSugerida", "Activa impresion de propina sugerida", IIf(Me.chkImprimePropinaSugerida.Value, "Verdadero", "Falso"), "lIntMercadoPagoQR", "Activa Mercado Pago Dinamico", IIf(Me.chkIntMercadoPagoQR.Value, "Verdadero", "Falso"), "lActivarAvisoRetenciones", "Activa Aviso Retenciones", IIf(Me.chklActivarAvisoRetenciones.Value, "Verdadero", "Falso"), "lActivaTresDecimales", "Activa Tres Decimales", IIf(Me.chklActivaTresDecimales.Value, "Verdadero", "Falso"), "lIntegracionCASSWII30CD", "Activa balanza CAS S2II30CD ", IIf(Me.chklIntegracionCASSWII30CD.Value, "Verdadero", "Falso"), _
                            "tDefaultOrigenVenta", "Origen de venta Predefinida", Me.cboOrigenVenta.BoundText, "lNoVisualizarItemCombos", "No Visualizar Item Combos Automaticos", IIf(Me.ChlNoVisualizarItemCombos.Value, "Verdadero", "Falso"), "lMostrarPrecioItemsCombo", "Visualizar Precios en el Desglose", IIf(Me.ChlMostrarPrecioItemsCombo.Value, "Verdadero", "Falso"), "lActivarMensajeCtsCorrientes", "Activar Mensaje en Cuentas Corrientes", Me.chklActivarMensajeCtsCorrientes.Value, "lIntMobileKDS", "Activar Integracion Mobile", Me.chklIntMobileKDS.Value, "lImprimirGuiaFormatoCorp", "Imprimir Formato de Guia 2", chkGuiaF2.Value, "lPermitirAnularMaxDias", "Permitir anular despues de los dias", IIf(Me.chklPermitirAnularMaxDias.Value, "Verdadero", "Falso"), "lOrdenarClientes", "Ordenar Lista Clientes, OTROS al final", IIf(Me.chklOrdenarClientes.Value, "Verdadero", "Falso"), _
                            "txtUserToken", "User Token", txtUserToken.Text, "txtApiKey", "ApiKey", txtApiKey.Text, "txtApiToken", "Api Token", txtApiToken.Text, "lVistaPreviaItemsCombo", "Vista Previa Items Combo", _
                            IIf(Me.chklIntegracionCASSWII30CD.Value, "Verdadero", "Falso"), "lImprimePropinaDoc", "Imprimir propina en documento", IIf(Me.chkimprimePropinaDoc.Value, "Verdadero", "Falso"), "IntFPAY", "Integracion FAPY", IIf(Me.chkIntFPAY.Value, "Verdadero", "Falso"), "lActivaCreditoCoorporativo", "Activa Credito Coorporativo", IIf(Me.ChklActivaCreditoCoorporativo.Value, "Verdadero", "Falso"), "IntegracionWebApi", "Activa Integracion Web Api", IIf(Me.chkIntegracionWebApi.Value, "Verdadero", "Falso"), "lActivaRestriClienteDelivery", "Activa Restri Cliente Delivery", IIf(Me.ChklActivaRestriClienteDelivery.Value, "Verdadero", "Falso"), "lMotivoAdm", "Flag Motivo Administrativo NC", IIf(chkMotivoAdm.Value, "Verdadero", "Falso"), "lBloqueoAnulaPedido", "Bloqueo de anulacion de pedidos", IIf(Me.ChkBloqueoAnulaPedido.Value, "Verdadero", "Falso"), "lBloqueoAnulaItemsPedido", "Bloqueo de anulacion de Items de pedido", IIf(Me.ChkBloqueoAnulaItemsPedido.Value, "Verdadero", "Falso"), _
                            "lOcultaRazonSocial", "Ocultar razon social de precuenta", IIf(Me.chkOcultarDesgloseRazonSocial.Value, "Verdadero", "Falso"), "lBloqueaCanal1", "Bloquea Canal 01", IIf(Me.ChkBloqueaCanal(1).Value, "Verdadero", "Falso"), "lBloqueaCanal2", "Bloquea Canal 02", IIf(Me.ChkBloqueaCanal(2).Value, "Verdadero", "Falso"), "lBloqueaCanal3", "Bloquea Canal 03", IIf(Me.ChkBloqueaCanal(3).Value, "Verdadero", "Falso"), "lBloqueaCanal4", "Bloquea Canal 04", IIf(Me.ChkBloqueaCanal(4).Value, "Verdadero", "Falso"), "lBloqueaCanal5", "Bloquea Canal 05", IIf(Me.ChkBloqueaCanal(5).Value, "Verdadero", "Falso"), "lImprimePrecuentaCabUnica", "Imprimir Cabecera Unica en Precuenta", IIf(chkPrecuentaCabUnica.Value, "Verdadero", "Falso"), "lBloquearPedidoSinMesa", "Bloquear pedidos sin mesa", IIf(ChkBloqueaPedidoSinMesa.Value, "Verdadero", "Falso"), "lAutorizacionSinMeza", "Autorización de opción sin mesa", IIf(chkAutorizacionSinMesa.Value, "Verdadero", "Falso"), _
                            "lSolicitaClaveEgreso", "Solicita clave recibo de egreso", IIf(Me.chkSolicitaClaveEgreso.Value, "Verdadero", "Falso"), "lBloqueoEdicionPedidoConsolaAgr", "Bloquea Edicion de pedido consola", IIf(Me.ChkBloqueoEdicionPedidoConsolaAgr.Value, "Verdadero", "false"), "lActivaPasswordMulticajero", "Activa Password Multicajero pago y documento", IIf(Me.ChkActivaPassworMulticajero.Value, "Verdadero", "Falso"), "lImprimePagoPrecuenta", "Imprime prepago en precuenta", IIf(Me.chkPrecuentaImprime(1).Value, "Verdadero", "Falso"), "lBusquedaReservas", "Activa Busqueda de Reservas", IIf(Me.chklBusquedaReserva, "Verdadero", "Falso"), "lActivarValesConsumo", "Activar Vales de Consumo", IIf(Me.chklActivarValesConsumo, "Verdadero", "Falso"))

                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                If lAuditoria = False Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
                End If
                                                
                  'GCAA 14102021
                                                   
                  'Cambiar el SQL
                  Isql = "insert into TCAJA( " & _
                  "tCaja, tDescripcion, tPrecuenta, vComanda, lComanda, lMotivoEliminaC, lMotivoElimina, lComboPrecuenta, lComboDocumento, lPasswordC, lPassword, tGrupo, lConsumo1, lConsumo2, lConsumo3, lPrecuenta, lAdicion, lPrecuentaAgrupada, tTipoPedido, lObliga, lMozo, lObligaPrinter, lObligaPrecuenta, lPax, lObligaCierre, lFiltroTipoPedido, nPuerto, tMensaje1, tMensaje2, lCancelacion, lDirecto, lCambioMesa, lVisaNet, lImpuestoPrecuenta, lDocumentoAgrupado, lOrden, lActivo, lValorCortesia, lObservacion, lCajaRapida, lPropiedadDocumento, lPropiedadPrecuenta, lPrecioNetoPrecuenta, nLimitePrecuenta, tUnidadNegocio, lPasswordTransferencia, nLimiteReimpresion, lCD, lMultiCajero, lFechaEntregaDelivery, lMCPV, lCCVOX, lMotorizado,lequivadolaprecuenta,tsubalmacen, lObservacionPrecuenta, lObservacionDocumento,lPasswordImportarPedido, lActivaImpDscAlternativa, lCompatibilidadTVS, nLongitudBarra,lpagorapido, lDisgrega,lpasswordporcobrar,lmodificatipopedido,TSUCURSAL, nBalanzaPuerto, lCapturaPeso, " & _
                  "lPagoRapidoPV, tTextoConsumo, lSiab, tSectorVenta,lcajamobile, lbloqueaprecuenta, lRotulado, lmultiAreaSubGrupo, lMultiAreaCaja, lHuella, lImprimeImagCabPrecuenta, lImprimeImagpiePrecuenta, laccesodespachopedido,lBuscaPedidoNumero,lCodigoReciboIngreso,lPagoRapidoMod,lWebAp,lMesa247,lConsumo4, lPrecuentaNoValorizada, LBuscarpedidovisualizargrilla ,lbuscarpedidofiltrarmesa, lClaveEnvioProduccion, EAN13,lObservacionCabDoc,tImpFiscal,lMotivoDescuento, lCajaContingencia, lImpPropina, lImpComandaf2, lPassOtrosPagos, lImpClienteCab, lBuscarPedidoBarra, lIntVisaNetQR, lImpComandaBarra, lbloqueadocconsumo,lCashDro, lIntUber, tIpCashDrow, lIntMercadoPagoQR,lIntPatio, lBloqueoAfectoInafecto,lIntMercadoPagoQREstatico,lIntPagoEfectivo, lPagoDirectoCD,lImprimePropinaSugerida,lActivarAvisoRetenciones, lActivaTresDecimales, lIntegracionCASSWII30CD, tDefaultOrigenVenta, lNoVisualizarItemCombos,lMostrarPrecioItemsCombo,lActivarMensajeCtsCorrientes,lIntMobileKDS,lImprimirGuiaFormatoCorp " & _
                   ", lPermitirAnularMaxDias, lOrdenarClientes,lUserToken,lApiKey,lApiToken, lVistaPreviaItemsCombo, lImprimePropinaDoc, IntFPAY, lActivaCreditoCoorporativo, IntegracionWebApi,lModoFacturacion,lActivarImpresionBlu, lActivaRestriClienteDelivery, lMotivoAdm,lBloqueoAnulaPedido,lBloqueoAnulaItemsPedido, lOcultaRazonSocial,lBloqueaCanal1, lBloqueaCanal2, lBloqueaCanal3, lBloqueaCanal4, lBloqueaCanal5, lImprimePrecuentaCabunica, lBloquearPedidoSinMesa, lAutorizacionSinMeza, lSolicitaClaveEgreso, lBloqueoEdicionPedidoConsolaAgr,lActivaPasswordMulticajero,lImprimePagoPrecuenta,lBusquedaReservas,lActivarValesConsumo) values ('" & txtCodigo.Text & "', " & _
                         " '" & txtDetallado.Text & "', " & _
                         " '" & cboPreCuenta.BoundText & "', " & _
                                chkVComanda.Value & ", " & chkComanda.Value & ", " & chkEliminaC.Value & ", " & chkElimina.Value & ", " & _
                                chkComboPrecuenta.Value & ", " & chkComboDocumento.Value & ", " & _
                                chkPasswordC.Value & ", " & _
                                chkPassword.Value & ", " & " '" & cboGrupo.BoundText & "', " & _
                                chkConsumo1.Value & ", " & chkConsumo2.Value & ", " & _
                                chkConsumo3.Value & ", " & _
                                chkPrecuenta.Value & ", " & _
                                chkAdicion.Value & ", " & chkAgrupada.Value & ", " & _
                         " '" & cboTipoPedido.BoundText & "',null, null,  " & _
                                chkObligaPrinter.Value & ", " & chkObligaPrecuenta.Value & ",null, " & _
                                chkObligaCierre.Value & ", " & chkFiltroTipoPedido.Value & ", " & _
                                val(txtPuerto.Text) & ", '" & txtMensaje1 & "', '" & txtMensaje1 & "', " & _
                                chkCancelacion.Value & ", " & chkDirecto.Value & ", " & chkCambioMesa.Value & ", " & chkVisaNet.Value & ", " & chkImpuestoPrecuenta.Value & "," & chkDocumentoAgrupado.Value & ", " & chkOrden.Value & ", " & chkActivo.Value & "," & chkValor.Value & ", " & _
                                chkObservacion.Value & ", " & chkCajaRapida.Value & ", " & chkPropiedadDocumento.Value & ", " & chkPropiedadPrecuenta.Value & ", " & chkPrecioNetoPrecuenta.Value & ", " & val(txtLimitePrecuenta.Text) & ", '" & cboUnidadNegocio.BoundText & "', " & val(txtLimitePrecuenta.Text) & ", " & chkPassword.Value & ", " & chkCD.Value & ", " & chkMultiCajero.Value & ", null, " & chkMCPV.Value & ", " & chkCCVOX.Value & ", null," & chkEquivaPrecuenta.Value & ",'" & IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText) & "'," & chkObservacionPrecuenta.Value & "," & chkObservacionDocumento.Value & "," & chkPasswordImportar.Value & "," & chkDescripcionAlternativa.Value & "," & chkCompatibilidadTVS.Value & ", '" & txtLongitudBarra.Text & "'," & chkPagoRapido.Value & "," & chkDisgrega.Value & "," & chkPasswordPorCobrar.Value & "," & chkModificaTipoPedido.Value & ", '" & _
                                cboSucursal.BoundText & "', " & val(txtBalanzaPuerto) & ", " & IIf(opcCapturaPeso.Value, 1, 0) & ", " & chkPagoRapidopv.Value & ",'" & txtTextoConsumo.Text & "', " & chkSiab.Value & ", '" & cboSectorVenta.BoundText & "', " & chkCajaMobile.Value & ", " & chkBloqueaPrecuenta.Value & "," & chkRotulado.Value & "," & chkMulti2.Value & ", " & chkMulti1.Value & ", " & chkHuella.Value & "," & chkImprimeImagCabPrecuenta.Value & "," & chkImprimeImagPiePrecuenta.Value & ", " & chkAccesoDespachoPedido.Value & "," & Me.chkBuscaPedido.Value & ", " & Me.chkCodigoReciboIngreso.Value & "," & chkPagoRapidoMod.Value & "," & chkWebAp.Value & "," & chkMesa247.Value & "," & chkConsumo4.Value & ", " & chkPrecuentaNoValorizada.Value & "," & Me.chkBuscaPedidoVisualizaGrilla.Value & "," & Me.chkBuscaPedidoFiltrarMesa.Value & ", " & Me.chkClaveEnvio.Value & " , " & IIf(chkEAN13.Value, 1, 0) & "," & _
                                IIf(chkObservacionCabDoc.Value, 1, 0) & ", '" & cboComprobante.BoundText & "', " & Me.chkMotDesc.Value & "," & Me.chkCajaContingencia.Value & ", " & Me.chkImpPropina.Value & ", " & Me.chkComandaF2.Value & ", " & Me.chkPassOtrosPagos.Value & ", " & Me.chkImpClienteCab.Value & ", " & Me.chkBuscarPedidoBarra.Value & ", " & Me.chkIntVisaNetQR.Value & ", " & Me.ChkImpComandaBarra.Value & ", " & ChkBloqEmiDoc.Value & "," & Me.ChkActivaCashDro.Value & "," & Me.chkActivaUber.Value & ",'" & Trim(Me.txtIpCashDrow.Text) & "'," & Me.chkIntMercadoPagoQR.Value & "," & Me.chkIntegracionPatio.Value & ", " & Me.ChkBloqAfectoInafecto.Value & "," & Me.chkIntMercadoPagoQREstatico.Value & "," & Me.chkPagoEfectivo.Value & "," & Me.ChkPagoDirectoCD.Value & " ," & _
                                Me.chkImprimePropinaSugerida.Value & ", " & Me.chklActivarAvisoRetenciones.Value & "," & Me.chklActivaTresDecimales.Value & "," & Me.chklIntegracionCASSWII30CD.Value & ", '" & Me.cboOrigenVenta.BoundText & "', " & Me.ChlNoVisualizarItemCombos.Value & ", " & Me.ChlMostrarPrecioItemsCombo.Value & " , " & Me.chklActivarMensajeCtsCorrientes.Value & ", " & Me.chklIntMobileKDS.Value & " , " & Me.chkGuiaF2.Value & " , " & Me.chklPermitirAnularMaxDias.Value & " ," & Me.chklOrdenarClientes.Value & "  ,'" & Me.txtUserToken.Text & "' ,'" & Me.txtApiKey.Text & "','" & Me.txtApiToken.Text & "'," & Me.ChkVistaPreviaItemsCombo.Value & "," & Me.chkimprimePropinaDoc.Value & "," & Me.chkIntFPAY.Value & "," & Me.ChklActivaCreditoCoorporativo.Value & " ," & Me.chkIntegracionWebApi.Value & "," & Me.chklModoFacturacion.Value & _
                                "," & Me.chklActivarImpresionBlu.Value & "," & Me.ChklActivaRestriClienteDelivery.Value & "," & Me.chkMotivoAdm.Value & "," & Me.ChkBloqueoAnulaPedido.Value & "," & Me.ChkBloqueoAnulaItemsPedido.Value & "," & Me.chkOcultarDesgloseRazonSocial.Value & "," & Me.ChkBloqueaCanal(1).Value & "," & Me.ChkBloqueaCanal(2).Value & "," & Me.ChkBloqueaCanal(3).Value & "," & Me.ChkBloqueaCanal(4).Value & "," & Me.ChkBloqueaCanal(5).Value & ", " & Me.chkPrecuentaCabUnica.Value & ", " & Me.ChkBloqueaPedidoSinMesa.Value & ", " & Me.chkAutorizacionSinMesa.Value & ", " & Me.chkSolicitaClaveEgreso.Value & ", " & Me.ChkBloqueoEdicionPedidoConsolaAgr.Value & ", " & Me.ChkActivaPassworMulticajero.Value & ", " & Me.chkPrecuentaImprime(1).Value & ", " & Me.chklBusquedaReserva.Value & ", " & Me.chklActivarValesConsumo.Value & ")"
                  Cn.Execute Isql

                  frmCaja.RsCabecera.Requery
                  frmCaja.RsCabecera.MoveLast
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
                               
                   RsGrilla.Requery
                   RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
                   RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"

                   If Not RsGrilla.EOF Then
                      RsGrilla.MoveFirst
                   End If
                   MsgBox "Registro Agregado", vbInformation, sMensaje
                       
               Else
                   sPasa = txtCodigo.Text
               
                   'Inserta Movimiento auditoria
                    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "02", sUsuario, sPasa, "", _
                          "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text, "tPrecuenta", "Impresora Precuenta", cboPreCuenta.BoundText, "lSiab", "Flag Activacion enlace Sistema SIAB", IIf(chkSiab.Value, "Verdadero", "Falso"), _
                          "tGrupo", "Grupo Predeterminado", cboGrupo.BoundText, "tTipoPedido", "Tipo Pedido Predeterminado", cboTipoPedido.BoundText, "tUnidadNegocio", "Unidad de Negocio", cboUnidadNegocio.BoundText, "tSucursal", "Sucursal", cboSucursal.BoundText, "tSubAlmacen", "Area de Produccion", IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), _
                          "nPuerto", "Visor Puerto", val(txtPuerto.Text), "tMensaje1", "Visor Mensaje 1", txtMensaje1.Text, "tMensaje2", "Visor Mensaje 2", txtMensaje2.Text, "nBalanzaPuerto", "Balanza Electronica Puerto", val(txtBalanzaPuerto.Text), "lVisaNet", "Flag Enlace VisaNet", IIf(chkVisaNet.Value, "Verdadero", "Falso"), "nLongitudBarra", "Lector Barra Longitud", txtLongitudBarra.Text, "lcapturaPeso", "Lector de Barra", IIf(opcCapturaPeso.Value, "Verdadero", "Falso"), _
                          "lComanda", "Flag Obligatoriedad Comanda Manual", IIf(chkComanda.Value, "Verdadero", "Falso"), "lObligaPrinter", "Flag Obliga Impresion Pedido", IIf(chkObligaPrinter.Value, "Verdadero", "Falso"), "lObligaPrecuenta", "Flag Obliga Impresion Precuenta", IIf(chkObligaPrecuenta.Value, "Verdadero", "Falso"), "lCancelacion", "Flag Obligatoriedad de Cancelacion", IIf(chkCancelacion.Value, "Verdadero", "Falso"), "lObservacion", "Flag Obligatoriedad Observacion", IIf(chkObservacion.Value, "Verdadero", "Falso"), "lConsumo1", "Flag Emision Rapida por Consumo", IIf(chkConsumo1.Value, "Verdadero", "Falso"), "lConsumo2", "Flag Emision por Consumo En Pagos y Division", IIf(chkConsumo2.Value, "Verdadero", "Falso"), "lConsumo3", "Flag Emision por Consumo en Caja Rapida", IIf(chkConsumo3.Value, "Verdadero", "Falso"), _
                          "lPrecuentaAgrupada", "Flag Impresion Agrupada en Precuenta", IIf(chkAgrupada.Value, "Verdadero", "Falso"), "lDocumentoAgrupado", "Flag Impresion Agrupada en Documentos", IIf(chkDocumentoAgrupado.Value, "Verdadero", "Falso"), "lComboPrecuenta", "Flag Impresion de Combos en Precuenta", IIf(chkComboPrecuenta.Value, "Verdadero", "Falso"), "lComboDocumento", "Flag Impresion de Combos en Documentos", IIf(chkComboDocumento.Value, "Verdadero", "Falso"), "lPropiedadPrecuenta", "Flag Impresion Propiedad en Precuenta", IIf(chkPropiedadPrecuenta.Value, "Verdadero", "Falso"), "lObservacionPrecuenta", "Flag Impresion de Observacion en Precuenta", IIf(chkObservacionPrecuenta.Value, "Verdadero", "Falso"), "lPropiedadDocumento", "Flag Impresion Propiedad en Documento", IIf(chkPropiedadDocumento.Value, "Verdadero", "Falso"), "lObservacionDocumento", "Flag Impresion Observacion en Documento", IIf(chkObservacionDocumento.Value, "Verdadero", "Falso"), _
                          "lPrecioNetoPrecuenta", "Flag Impresion Prec Neto en Precuenta", IIf(chkPrecioNetoPrecuenta.Value, "Verdadero", "Falso"), "lPrecuenta", "Flag Permite Cambiar Impresora Precuenta", IIf(chkPrecuenta.Value, "Verdadero", "Falso"), "lImpuestoPrecuenta", "Flag Impresion Impuesto Desglos. Precuenta", IIf(chkImpuestoPrecuenta.Value, "Verdadero", "Falso"), "lCambioMesa", "Flag Impresion de Cambio de Mesa", IIf(chkCambioMesa.Value, "Verdadero", "Falso"), "lValorCortesia", "Flag Impresion Valorizada de Cortesias", IIf(chkValor.Value, "Verdadero", "Falso"), "lequivadolaprecuenta", "Flag Impresion de Equivalencia Dolares en Precuenta", IIf(chkEquivaPrecuenta.Value, "Verdadero", "Falso"), "lActivaImpDscAlternativa", "Flag Impresion Descripcion Alternativa", IIf(chkDescripcionAlternativa.Value, "Verdadero", "Falso"), "nLimitePrecuenta", "Limite de Precuentas", val(txtLimitePrecuenta.Text), "nLimiteReimpresion", "Limite de Re Impresiones Pedido", val(txtLimiteReimpresion.Text), _
                          "vComanda", "Flag Activa Ingreso Comanda Manual", IIf(chkVComanda.Value, "Verdadero", "Falso"), "lMotivoEliminaC", "Flag Pide Motivo Elimina Pedido", IIf(chkEliminaC.Value, "Verdadero", "Falso"), "lPasswordC", "Flag Activa Password Eliminacion Pedido", IIf(chkPasswordC.Value, "Verdadero", "Falso"), "lMotivoElimina", "Flag Pide Motivo Elimina Producto", IIf(chkElimina.Value, "Verdadero", "Falso"), "lPassword", "Flag Activa Password Elimina Producto", IIf(chkPassword.Value, "Verdadero", "Falso"), "lObligacierre", "Flag Activa Password Cierre Turno", IIf(chkObligaCierre.Value, "Verdadero", "Falso"), "lPasswordTransferencia", "Flag Activa Password Transferencia", IIf(chkPasswordTransferencia.Value, "Verdadero", "Falso"), "lpasswordporcobrar", "Flag Activa Password Por Cobrar", IIf(chkPasswordPorCobrar.Value, "Verdadero", "Falso"), "lPasswordImportarPedido", "Flag Activa Password Importar Pedido", IIf(chkPasswordImportar.Value, "Verdadero", "Falso"), _
                          "lFiltroTipoPedido", "Flag Permite Importar Pedidos por Canal", IIf(chkFiltroTipoPedido.Value, "Verdadero", "Falso"), "lAdicion", "Flag Permite Transferencias", IIf(chkAdicion.Value, "Verdadero", "Falso"), "lmodificatipopedido", "Flag Permite Modificar Tipo de Pedido", IIf(chkModificaTipoPedido.Value, "Verdadero", "Falso"), "lCajaRapida", "Flag Ingreso Directo a Caja Rapida", IIf(chkCajaRapida.Value, "Verdadero", "Falso"), "lPagoRapido", "Flag Ingreso a Pago Rapido desde Caja Rapida", IIf(chkPagoRapido.Value, "Verdadero", "Falso"), "lOrden", "Flag Activa Control Enum. Automatica", IIf(chkOrden.Value, "Verdadero", "Falso"), "lDirecto", "Flag Activa Control de Envios Directos", IIf(chkDirecto.Value, "Verdadero", "Falso"), "lDisgrega", "Flag Disgregar en Dos Partes", IIf(chkDisgrega.Value, "Verdadero", "Falso"), "lCD", "Flag Activa Caja Central Delivery", IIf(chkCD.Value, "Verdadero", "Falso"), "lCCVOX", "Flag Activa Caja Delivery CCVOX", IIf(chkCCVOX.Value, "Verdadero", "Falso"), _
                          "lMultiCajero", "Flag Activa Multicajero Caja Rapida", IIf(chkMultiCajero.Value, "Verdadero", "Falso"), "lMCPV", "Flag Activa Multicajero Salon", IIf(chkMCPV.Value, "Verdadero", "Falso"), "lCompatibilidadTVS", "Flag Permite Compatibilidad con TVS", IIf(chkCompatibilidadTVS.Value, "Verdadero", "Falso"), "lPagoRapidoPV", "Flag Ingreso a Pago Rapido desde Punto Venta", IIf(chkPagoRapidopv.Value, "Verdadero", "Falso"), "tTextoConsumo", "Motivo de Consumo Predeterminado", txtTextoConsumo.Text, "tSectorVenta", "SectorVenta", cboSectorVenta.BoundText, "lCajaMobile", "Flag Caja Mobile", IIf(chkCajaMobile.Value, "Verdadero", "Falso"), "lBloqueaPrecuenta", "Bloquea Precuenta", IIf(chkBloqueaPrecuenta.Value, "Verdadero", "Falso"), "lRotulado", "Enlace Rotulado", IIf(chkRotulado.Value, "Verdadero", "Falso"), "lMultiAreaSubGrupo", "Flag Multi Area Por SubGrupo", IIf(Me.chkMulti2.Value, "Verdadero", "Falso"), _
                          "lMultiAreaCaja", "Flag Multi Area ", IIf(Me.chkMulti1.Value, "Verdadero", "Falso"), "lHuella", "Flag Huella ", IIf(Me.chkHuella.Value, "Verdadero", "Falso"), "lImprimeImagCabPrecuenta", "Imagen Cabecera Precuenta", IIf(Me.chkImprimeImagCabPrecuenta.Value, "Verdadero", "Falso"), "lImprimeImagpiePrecuenta", "Imagen Pie Precuenta", IIf(Me.chkImprimeImagPiePrecuenta.Value, "Verdadero", "Falso"), "lAccesoDespachoPedido", "Acceso Despacho Pedido", IIf(Me.chkAccesoDespachoPedido.Value, "Verdadero", "Falso"), "LBuscarpedidovisualizargrilla", "Buscar Pedido Visualizar Grilla", IIf(Me.chkBuscaPedidoVisualizaGrilla.Value, "Verdadero", "Falso"), "lbuscarpedidofiltrarmesa", "Buscar Pedido Filtrar Mesa", IIf(Me.chkBuscaPedidoFiltrarMesa.Value, "Verdadero", "Falso"), "lMotivoDescuento", "Imprime Motivo Descuento", IIf(Me.chkMotDesc.Value, "Verdadero", "Falso"), "lCajaContingencia", "Activa Caja Contingencia", IIf(Me.chkCajaContingencia.Value, "Verdadero", "Falso"), _
                          "lImpPropina", "Solicita Propina Imp Prec", IIf(Me.chkImpPropina.Value, "Verdadero", "Falso"), "lImpComandaf2", "imprime comanda formato 2", IIf(Me.chkComandaF2.Value, "Verdadero", "Falso"), "lPassOtrosPagos", "Activa Password otros pagos", IIf(Me.chkPassOtrosPagos.Value, "Verdadero", "Falso"), "lImpClienteCab", "Imprime Cliente en Cabecera", IIf(Me.chkImpClienteCab.Value, "Verdadero", "Falso"), "lBuscarPedidoBarra", "Buscar pedido no enviados por codigo de barra", IIf(chkBuscarPedidoBarra.Value, "Verdadero", "Falso"), "lIntVisaNetQR", "IntegracionVisaNet QR", IIf(Me.chkIntVisaNetQR.Value, "Verdadero", "Falso"), "lImpComandaBarra", "Imprime comanda con barra por Item", IIf(Me.ChkImpComandaBarra.Value, "Verdadero", "Falso"), "lIntUber", "Activa Integracion Uber", IIf(Me.chkActivaUber.Value, "Verdadero", "Falso"), "tIpCashDrow", "Ip para la conexion a Cash Drow", Me.txtIpCashDrow.Text, "lBloqueoAfectoInafecto", "Activa el proceso de bloqueo de afectos e inafectos en documentos", _
                          IIf(Me.ChkBloqAfectoInafecto.Value, "Verdadero", "Falso"), "lIntMercadoPagoQR", "Activa Mercado Pago Estatico", IIf(Me.chkIntMercadoPagoQREstatico.Value, "Verdadero", "Falso"), "lIntPagoEfectivo", "INtegracion pago Efectivo", IIf(Me.chkPagoEfectivo.Value, "Verdadero", "Falso"), "lPagoDirectoCD", "Pago Directo Central Delivery", IIf(Me.ChkPagoDirectoCD.Value, "Verdadero", "Falso"), "lImprimePropinaSugerida", "Activa impresion de propina sugerida", IIf(Me.chkImprimePropinaSugerida.Value, "Verdadero", "Falso"), "lIntMercadoPagoQR", "Activa Mercado Pago Dinamico", IIf(Me.chkIntMercadoPagoQR.Value, "Verdadero", "Falso"), "lActivarAvisoRetenciones", "Activa Aviso Retenciones", IIf(Me.chklActivarAvisoRetenciones.Value, "Verdadero", "Falso"), "lActivaTresDecimales", "Activa Tres Decimales", IIf(Me.chklActivaTresDecimales.Value, "Verdadero", "Falso"), "lIntegracionCASSWII30CD", "Activa balanza CAS S2II30CD ", IIf(Me.chklIntegracionCASSWII30CD.Value, "Verdadero", "Falso"), _
                          "tDefaultOrigenVenta", "Origen de venta Predefinida", Me.cboOrigenVenta.BoundText, "lNoVisualizarItemCombos", "No Visualizar Item Combos Automaticos", IIf(Me.ChlNoVisualizarItemCombos.Value, "Verdadero", "Falso"), "lMostrarPrecioItemsCombo", "Visualizar Precios en el Desglose", IIf(Me.ChlMostrarPrecioItemsCombo.Value, "Verdadero", "Falso"), "lActivarMensajeCtsCorrientes", "Activar Mensaje en Cuentas Corrientes", IIf(Me.chklActivarMensajeCtsCorrientes.Value, "Verdadero", "Falso"), "lIntMobileKDS", "Activar Integracion Mobile", IIf(Me.chklIntMobileKDS.Value, "Verdadero", "Falso"), "lImprimirGuiaFormatoCorp", "Imprimir Formato de Guia 2", IIf(chkGuiaF2.Value, "Verdadero", "Falso"), "lPermitirAnularMaxDias", "Permitir anular despues de los dias", IIf(Me.chklPermitirAnularMaxDias.Value, "Verdadero", "Falso"), "lOrdenarClientes", "Ordenar Lista Clientes, OTROS al final", IIf(Me.chklOrdenarClientes.Value, "Verdadero", "Falso"), "lUserToken", "User Token", txtUserToken.Text, "lApiKey", _
                          "ApiKey", txtApiKey.Text, "lApiToken", "Api Token", txtApiToken.Text, "lVistaPreviaItemsCombo", "Vista Previa Items Combo ", IIf(Me.ChkVistaPreviaItemsCombo.Value, "Verdadero", "Falso"), _
                          "lImprimePropinaDoc", "Imprimir propina en documento", IIf(Me.chkimprimePropinaDoc.Value, "Verdadero", "Falso"), "IntFPAY", "Integracion FPAY", IIf(Me.chkIntFPAY.Value, "Verdadero", "Falso"), "lActivaCreditoCoorporativo", "Activa Credito Coorporativo", IIf(Me.ChklActivaCreditoCoorporativo.Value, "Verdadero", "Falso"), "IntegracionWebApi", "Activa Integracion Web Api", IIf(Me.chkIntegracionWebApi.Value, "Verdadero", "Falso"), "lActivaRestriClienteDelivery", "Activa Restri Cliente Delivery", IIf(Me.ChklActivaRestriClienteDelivery.Value, "Verdadero", "Falso"), "lMotivoAdm", "Flag Motivo Administrativo NC", IIf(chkMotivoAdm.Value, "Verdadero", "Falso"), "lBloqueoAnulaPedido", "Bloqueo de anulacion de pedidos", IIf(Me.ChkBloqueoAnulaPedido.Value, "Verdadero", "Falso"), "lBloqueoAnulaItemsPedido", "Bloqueo de anulacion de Items de pedido", IIf(Me.ChkBloqueoAnulaItemsPedido.Value, "Verdadero", "Falso"), _
                          "lOcultaRazonSocial", "Ocultar razon social de precuenta", IIf(Me.chkOcultarDesgloseRazonSocial.Value, "Verdadero", "Falso"), "lBloqueaCanal1", "Bloquea Canal 01", IIf(Me.ChkBloqueaCanal(1).Value, "Verdadero", "Falso"), "lBloqueaCanal2", "Bloquea Canal 02", IIf(Me.ChkBloqueaCanal(2).Value, "Verdadero", "Falso"), "lBloqueaCanal3", "Bloquea Canal 03", IIf(Me.ChkBloqueaCanal(3).Value, "Verdadero", "Falso"), "lBloqueaCanal4", "Bloquea Canal 04", IIf(Me.ChkBloqueaCanal(4).Value, "Verdadero", "Falso"), "lBloqueaCanal5", "Bloquea Canal 05", IIf(Me.ChkBloqueaCanal(5).Value, "Verdadero", "Falso"), "lImprimePrecuentaCabUnica", "Imprimir Cabecera Unica en Precuenta", IIf(chkPrecuentaCabUnica.Value, "Verdadero", "Falso"), "lBloquearPedidoSinMesa", "Bloquear pedidos sin mesa", IIf(ChkBloqueaPedidoSinMesa.Value, "Verdadero", "Falso"), "lAutorizacionSinMeza", "Autorización de opción sin mesa", IIf(chkAutorizacionSinMesa.Value, "Verdadero", "Falso"), _
                          "lSolicitaClaveEgreso", "Solicita clave recibo de egreso", IIf(Me.chkSolicitaClaveEgreso.Value, "Verdadero", "Falso"), "lBloqueoEdicionPedidoConsolaAgr", "Bloquea Edicion de pedido consola", IIf(Me.ChkBloqueoEdicionPedidoConsolaAgr.Value, "Verdadero", "false"), "lActivaPasswordMulticajero", "Activa Password Multicajero pago y documento", IIf(Me.ChkActivaPassworMulticajero.Value, "Verdadero", "Falso"), "lImprimePagoPrecuenta", "Imprime prepago en precuenta", IIf(Me.chkPrecuentaImprime(1).Value, "Verdadero", "Falso"), "lBusquedaReservas", "Activa Busqueda de Reservas", IIf(Me.chklBusquedaReserva.Value, "Verdadero", "Falso"), "lActivarValesConsumo", "Activar Vales de Consumo", IIf(Me.chklActivarValesConsumo.Value, "Verdadero", "Falso"))

                   If lAuditoria = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                
                  'Cambiar el SQL
                  Isql = "update TCAJA set " & _
                         "tDescripcion ='" & txtDetallado.Text & "', tPrecuenta ='" & cboPreCuenta.BoundText & "', vComanda =" & chkVComanda.Value & ", " & _
                         "lComanda =" & chkComanda.Value & ", lMotivoEliminaC =" & chkEliminaC.Value & ", lMotivoElimina =" & chkElimina.Value & ", " & _
                         "lComboPrecuenta =" & chkComboPrecuenta.Value & ", lComboDocumento =" & chkComboDocumento.Value & ", lPasswordC =" & chkPasswordC.Value & ", " & _
                         "lPassword =" & chkPassword.Value & ", lPasswordTransferencia =" & chkPasswordTransferencia.Value & ", " & _
                         "tGrupo ='" & cboGrupo.BoundText & "', tUnidadNegocio ='" & cboUnidadNegocio.BoundText & "',tsubalmacen ='" & IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText) & "', " & _
                         "lConsumo1 =" & chkConsumo1.Value & ", lConsumo2 =" & chkConsumo2.Value & ", lConsumo3 =" & chkConsumo3.Value & ", " & _
                         "lPrecuenta =" & chkPrecuenta.Value & ", lDisgrega =" & chkDisgrega.Value & "," & "lActivaImpDscAlternativa =" & chkDescripcionAlternativa.Value & ",lcajamobile=" & chkCajaMobile.Value & ",  " & _
                         "lAdicion =" & chkAdicion.Value & ", lCCVOX =" & chkCCVOX.Value & ", lbloqueaprecuenta=" & chkBloqueaPrecuenta.Value & ", " & _
                         "lPrecuentaAgrupada =" & chkAgrupada.Value & ", lObservacionPrecuenta =" & chkObservacionPrecuenta.Value & ",lObservacionDocumento =" & chkObservacionDocumento.Value & ",lBuscaPedidoNumero=" & Me.chkBuscaPedido.Value & ",  " & _
                         "lObliga =  0 , lMozo =  0 ,lMotorizado =  0 ,lObligaPrinter =" & IIf(chkObligaPrinter.Value, 1, 0) & ", lObligaPrecuenta =" & IIf(chkObligaPrecuenta.Value, 1, 0) & ", lCodigoReciboIngreso=" & Me.chkCodigoReciboIngreso.Value & ", " & _
                         "tTipoPedido ='" & cboTipoPedido.BoundText & "', lPax =0 ,lMultiAreaSubGrupo=" & Me.chkMulti2.Value & ", lmultiAreaCaja=" & Me.chkMulti1.Value & ",  " & _
                         "lObligacierre =" & chkObligaCierre.Value & ", lCancelacion =" & chkCancelacion.Value & ", lDirecto =" & chkDirecto.Value & ", lDocumentoAgrupado =" & chkDocumentoAgrupado.Value & ", " & _
                         "lFiltroTipoPedido=" & chkFiltroTipoPedido.Value & ", lMCPV =" & chkMCPV.Value & ", lImprimeImagCabPrecuenta=" & chkImprimeImagCabPrecuenta.Value & ", lImprimeImagpiePrecuenta=" & chkImprimeImagPiePrecuenta.Value & ",  " & _
                         "lActivo =" & chkActivo.Value & ", lCambioMesa =" & chkCambioMesa.Value & ", lVisaNet =" & chkVisaNet.Value & ", lImpuestoPrecuenta =" & chkImpuestoPrecuenta.Value & ", " & _
                         "nPuerto =" & val(txtPuerto.Text) & ", tMensaje1='" & txtMensaje1.Text & "', tMensaje2='" & txtMensaje2.Text & "', lOrden= " & chkOrden.Value & ", lValorCortesia=" & chkValor.Value & ", nLongitudBarra ='" & txtLongitudBarra.Text & "', " & _
                         "lObservacion =" & chkObservacion.Value & ", lCajaRapida =" & chkCajaRapida.Value & ", lPropiedadDocumento =" & chkPropiedadDocumento.Value & ", lPrecioNetoPrecuenta =" & chkPrecioNetoPrecuenta.Value & ", lPropiedadPrecuenta =" & chkPropiedadPrecuenta.Value & ", " & "nLimitePrecuenta=" & val(txtLimitePrecuenta.Text) & ", nLimiteReimpresion=" & val(txtLimiteReimpresion.Text) & ",lCD =" & chkCD.Value & ", lMultiCajero =" & chkMultiCajero.Value & ", lFechaEntregaDelivery =0,lequivadolaprecuenta =" & chkEquivaPrecuenta.Value & ",lPasswordImportarPedido =" & chkPasswordImportar.Value & ", lCompatibilidadTVS =" & chkCompatibilidadTVS.Value & ", lPagoRapido =" & chkPagoRapido.Value & ", lpasswordporcobrar =" & chkPasswordPorCobrar.Value & ", lmodificatipopedido =" & chkModificaTipoPedido.Value & ", tSucursal ='" & cboSucursal.BoundText & "', nBalanzaPuerto= " & val(txtBalanzaPuerto.Text) & ", " & _
                         "lCapturaPeso = " & IIf(opcCapturaPeso.Value, 1, 0) & ", laccesodespachopedido=" & chkAccesoDespachoPedido.Value & ", lPagoRapidoPV = " & chkPagoRapidopv.Value & ", LBuscarpedidovisualizargrilla=" & Me.chkBuscaPedidoVisualizaGrilla.Value & ", lbuscarpedidofiltrarmesa=" & Me.chkBuscaPedidoFiltrarMesa.Value & " ,  " & _
                         "tTextoConsumo = '" & txtTextoConsumo.Text & "', lSiab =" & chkSiab.Value & ",lRotulado=" & chkRotulado.Value & ", tSectorVenta ='" & cboSectorVenta.BoundText & "', " & _
                         "lPagoRapidoMod = " & chkPagoRapidoMod.Value & ", lClaveEnvioProduccion = " & chkClaveEnvio.Value & ", EAN13 = " & chkEAN13.Value & ", lWebAp = " & chkWebAp.Value & ",lMesa247 = " & chkMesa247.Value & ", lConsumo4 = " & chkConsumo4.Value & ", lPrecuentaNoValorizada = " & chkPrecuentaNoValorizada.Value & ", " & _
                         "lHuella = " & chkHuella.Value & ", lObservacionCabDoc = " & IIf(chkObservacionCabDoc.Value, 1, 0) & ", tImpFiscal='" & cboComprobante.BoundText & "', lMotivoDescuento=" & Me.chkMotDesc.Value & ", lCajaContingencia= " & Me.chkCajaContingencia.Value & ", lImpPropina= " & Me.chkImpPropina.Value & ", lImpComandaf2=" & Me.chkComandaF2.Value & ", lPassOtrosPagos = " & Me.chkPassOtrosPagos.Value & ", lImpClienteCab=" & Me.chkImpClienteCab.Value & ", lBuscarPedidoBarra= " & Me.chkBuscarPedidoBarra.Value & ", lIntVisaNetQR=" & Me.chkIntVisaNetQR.Value & " , lImpComandaBarra= " & Me.ChkImpComandaBarra.Value & ", lbloqueadocconsumo = " & Me.ChkBloqEmiDoc.Value & ", lCashDro=" & Me.ChkActivaCashDro.Value & ", lIntUber=" & Me.chkActivaUber.Value & ", tIpCashDrow='" & Trim(Me.txtIpCashDrow.Text) & "', lIntMercadoPagoQR=" & Me.chkIntMercadoPagoQR.Value & ", " & _
                         " lIntPatio=" & Me.chkIntegracionPatio.Value & ", lBloqueoAfectoInafecto=" & Me.ChkBloqAfectoInafecto.Value & ", lIntMercadoPagoQREstatico=" & Me.chkIntMercadoPagoQREstatico.Value & ", lIntPagoEfectivo=" & Me.chkPagoEfectivo.Value & ", lPagoDirectoCD =" & Me.ChkPagoDirectoCD.Value & ", lImprimePropinaSugerida=" & Me.chkImprimePropinaSugerida.Value & " , lActivarAvisoRetenciones=" & Me.chklActivarAvisoRetenciones.Value & ", " & _
                         " lActivaTresDecimales=" & Me.chklActivaTresDecimales.Value & ", lIntegracionCASSWII30CD=" & Me.chklIntegracionCASSWII30CD.Value & ", tDefaultorigenVenta='" & Me.cboOrigenVenta.BoundText & "', lNoVisualizarItemCombos = " & Me.ChlNoVisualizarItemCombos.Value & " , lMostrarPrecioItemsCombo = " & Me.ChlMostrarPrecioItemsCombo.Value & " , lActivarMensajeCtsCorrientes = " & Me.chklActivarMensajeCtsCorrientes.Value & "  , lIntMobileKDS = " & Me.chklIntMobileKDS.Value & " , lImprimirGuiaFormatoCorp = " & Me.chkGuiaF2.Value & ", lPermitirAnularMaxDias = " & Me.chklPermitirAnularMaxDias.Value & ", lOrdenarClientes = " & Me.chklOrdenarClientes.Value & ", lUserToken = '" & Me.txtUserToken.Text & "', lApiKey = '" & Me.txtApiKey.Text & "', lApiToken = '" & Me.txtApiToken.Text & "', lVistaPreviaItemsCombo = " & Me.ChkVistaPreviaItemsCombo.Value & " , lImprimePropinaDoc = " & Me.chkimprimePropinaDoc & ", lActivaCreditoCoorporativo = " & Me.ChklActivaCreditoCoorporativo.Value & "" & _
                         " , IntegracionWebApi = " & Me.chkIntegracionWebApi.Value & ", lModoFacturacion = " & Me.chklModoFacturacion.Value & ", lActivarImpresionBlu = " & Me.chklActivarImpresionBlu.Value & ", lActivaRestriClienteDelivery = " & Me.ChklActivaRestriClienteDelivery.Value & ", lMotivoAdm = " & Me.chkMotivoAdm.Value & " , lBloqueoAnulaPedido = " & Me.ChkBloqueoAnulaPedido.Value & ", lBloqueoAnulaItemsPedido = " & Me.ChkBloqueoAnulaItemsPedido.Value & ", lOcultaRazonSocial = " & Me.chkOcultarDesgloseRazonSocial.Value & _
                         " , lBloqueaCanal1 = " & Me.ChkBloqueaCanal(1).Value & ", lBloqueaCanal2= " & Me.ChkBloqueaCanal(2).Value & ", lBloqueaCanal3 = " & Me.ChkBloqueaCanal(3).Value & ", lBloqueaCanal4 = " & Me.ChkBloqueaCanal(4).Value & ", lBloqueaCanal5 = " & Me.ChkBloqueaCanal(5).Value & ", lImprimePrecuentaCabunica = " & Me.chkPrecuentaCabUnica.Value & ", lBloquearPedidoSinMesa = " & Me.ChkBloqueaPedidoSinMesa.Value & ", lAutorizacionSinMeza = " & Me.chkAutorizacionSinMesa.Value & ", lSolicitaClaveEgreso = " & Me.chkSolicitaClaveEgreso.Value & ", lBloqueoEdicionPedidoConsolaAgr=" & Me.ChkBloqueoEdicionPedidoConsolaAgr.Value & ", lActivaPasswordMulticajero = " & Me.ChkActivaPassworMulticajero.Value & ", lImprimePagoPrecuenta=" & Me.chkPrecuentaImprime(1).Value & ", lBusquedaReservas=" & Me.chklBusquedaReserva.Value & ", lActivarValesConsumo=" & Me.chklActivarValesConsumo.Value & "  where tCaja = '" & txtCodigo & "'"
                                                
                   Cn.Execute Isql
                   
                   If sCaja = txtCodigo Then
                    lIntPatio = Me.chkIntegracionPatio.Value
                   End If
                   
                   nPos = frmCaja.RsCabecera.Bookmark
                   frmCaja.RsCabecera.Requery
                   If frmCaja.RsCabecera.RecordCount = 0 Then
                      frmCaja.RsCabecera.Filter = adFilterNone
                   End If
                   frmCaja.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               cmdOpcionGrilla(0).Enabled = True
               cmdOpcionGrilla(1).Enabled = True
               cmdOpcionGrilla(2).Enabled = True
               
               
               'Cambiar el SQL
                  Isql = "update TCAJACANALVENTA set " & _
                         "tcanalventa1 = " & ChCanal(1).Value & " ," & _
                         "tcanalventa2 = " & ChCanal(2).Value & " ," & _
                         "tcanalventa3 = " & ChCanal(3).Value & " ," & _
                         "tcanalventa4 = " & ChCanal(4).Value & " ," & _
                         "tcanalventa5 = " & ChCanal(5).Value & "  " & _
                         " where tCaja = '" & txtCodigo.Text & "'  "
                       
                   Cn.Execute Isql

          Case Is = 2 ' Eliminar
               
               If frmCaja.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
            
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Caja " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If txtCodigo.Text = sCaja Then
                  MsgBox "No se puede eliminar la Caja activa", vbCritical, sMensaje
                  Exit Sub
               End If
               
               sPasa = txtCodigo.Text
               
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "03", sUsuario, sPasa, "", _
                            "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text)
               
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                                  
               
               'Cambia el Delete
               Cn.Execute "update tcaja set lactivo=0 where tCaja = '" & txtCodigo.Text & "'"
               'Cn.Execute "delete from TCAJA where tCaja = '" & txtCodigo & "'"
               'Cn.Execute "delete from TTIPODOCUMENTOIMPRESORA where tCaja ='" & txtCodigo.Text & "'"
               frmCaja.RsCabecera.Requery
               
               If frmCaja.RsCabecera.RecordCount <> 0 Then
                  frmCaja.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
          Case Is = 4
            Dim xCriterio As String
            Isql = "select t.Descripcion as Canal ,o.CodOrigenVenta as Codigo, o.Descripcion  from vOrigenVenta o left join vTipoPedido t on  o.CodCanalVenta = t.Codigo where o.CodOrigenVenta not in (select CodigoOrigenventa from TCAJAORIGEN_BLOQUEO where tCaja='" & txtCodigo.Text & "') order by t.Codigo"
            
            frmBusca.cboCriterio.Enabled = True
            frmBusca.nPredeterm = 2
            Call ConfGrilla(3, frmBusca.grdGrilla, "Canal", 2, "Canal", 1800, 0, 0, "", _
                                                   "Codigo", 2, "Codigo", 1500, 0, 0, "", _
                                                   "Descripcion", 2, "Descripcion", 3500, 0, 0, "")
            frmBusca.Show vbModal
            If Not wEnter Then
               Exit Sub
            Else
                Cn.Execute ("if not exists (select * from TCAJAORIGEN_BLOQUEO where tcaja='" & Me.txtCodigo.Text & "' and CodigoOrigenventa='" & sCodigo & "'  ) begin insert into TCAJAORIGEN_BLOQUEO (tCaja,CodigoOrigenVenta) values ('" & Me.txtCodigo.Text & "','" & sCodigo & "') end")
                RsBloqueoOrigenVenta.Requery
            End If
           
            
   End Select
End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               'Cambiar los Controles
                With RsGrilla
                     'Cuadro de Texto
                     cboTipoDocumento.Text = ""
                     cboImpresora.Text = ""
                     cboFormulario.Text = ""
                     txtSerie.Text = ""
                     txtSerie2.Text = ""
                     txtCorrelativo.Text = ""
                     txtCorrelativo2.Text = ""
                     txtDescripcion.Text = ""
                     txtAutorizacion.Text = ""
                     chLImprimeImageCab.Value = 0
                     chLImprimeImagePie.Value = 0
                     chkFacturacionE.Value = 0
                     chkFacturacionOfisis.Value = 0
                     txtPrefijoEnlace.Text = ""
                End With
                SubDetalle False
                wAgrega = True
                cboTipoDocumento.Enabled = True
          
          Case Is = 1 ' Modificar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetalle False
               wAgrega = False
               SubAsignar
               cboTipoDocumento.Enabled = False
          
          Case Is = 2 ' Eliminar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubAsignar
               If MsgBox("Seguro de Eliminar este Tipo Documento " & cboTipoDocumento.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                   sPasa = txtCodigo.Text
                   
                   If pais = "002" Then 'Ecuador
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "03", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                      "tUsuario", "Usuario", sUsuario)
                   Else
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "03", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                      "tUsuario", "Usuario", sUsuario)
                   End If
          
                   If lAuditoria = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                   
               'Cambia el Delete
               Cn.Execute "delete from TTIPODOCUMENTOIMPRESORA where tCaja ='" & txtCodigo.Text & "' and tTipoEmision ='" & cboTipoDocumento.BoundText & "'"
               RsGrilla.Requery
               If RsGrilla.RecordCount <> 0 Then
                  RsGrilla.MoveLast
               End If
          
          Case Is = 3 ' Grabar
          
               If pais = "000" Then
                     If Len(Trim(txtSerie.Text)) <> 5 Then
                          MsgBox "El número de serie debe ser de 5 caracteres", vbCritical, sMensaje
                          Exit Sub
                     End If
               End If
          
               If wAgrega Then
                   RsGrilla.Find ("tTipoEmision ='" & cboTipoDocumento.BoundText & "'")
                   If Not RsGrilla.EOF Then
                      MsgBox "Tipo de Documento " & cboTipoDocumento.Text & " ya ingresado", vbCritical, sMensaje
                      Exit Sub
                   End If
                          
                   sPasa = txtCodigo.Text
                   
                   'JAPN 03012024 chkSolicitarConfirmacion para confirmar la generaciòn de documentos
                   
                   If pais = "002" Then 'Ecuador
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "01", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", cboLocal.Text & txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                      "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.Value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), _
                      "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.Value, "Verdadero", "Falso"), _
                      "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.Value, "Verdadero", "Falso"), _
                      "lConfirmacionDoc", "Confirma Generacion de Documentos", IIf(chkSolicitarConfirmacion.Value, "Verdadero", "Falso"), _
                      "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.Value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.Value, "Verdadero", "Falso"), "lImpDocMayorCero", "Impresion de Platos Mayor a Cero", IIf(Me.chkMayorCero.Value, "Verdadero", "Falso"))
                      
                   Else
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "01", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                      "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.Value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), _
                      "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.Value, "Verdadero", "Falso"), _
                      "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.Value, "Verdadero", "Falso"), _
                      "lConfirmacionDoc", "Confirma Generacion de Documentos", IIf(chkSolicitarConfirmacion.Value, "Verdadero", "Falso"), _
                      "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.Value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.Value, "Verdadero", "Falso"), "lImpDocMayorCero", "Impresion de Platos Mayor a Cero", IIf(Me.chkMayorCero.Value, "Verdadero", "Falso"))
                   End If
                   
                   If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                   End If
                   If pais = "002" Then 'Ecuador
                            Isql = "insert into TTIPODOCUMENTOIMPRESORA( " & _
                                   "tCaja, tTipoEmision, tImpresora, tDescripcion, tFormulario, tSerie, tUltimoNumero, tUsuario, lResumen, lImpuesto1, lImpuesto2, lImpuesto3, fRegistro,lEquivaDolares, lImprimeImageCab,lImprimeImagePie, lFacturacionElectronica,tNumeroAutorizacion,fInicio,fCaducidad,tPrefijoEnlace,lDocumentoElectronicoOfisis,limprimeresumen,lOpGravInaf, lImpProdDesc, limpDocMayorCero, lConfirmacionDoc) " & _
                                   "values (  '" & txtCodigo.Text & "', " & _
                                            " '" & cboTipoDocumento.BoundText & "', " & _
                                            " '" & cboImpresora.BoundText & "', " & _
                                            " '" & txtDescripcion.Text & "', " & _
                                            " '" & cboFormulario.BoundText & "', " & _
                                            " '" & cboLocal.Text & txtSerie2.Text & "', " & _
                                            " '" & txtCorrelativo2.Text & "', " & _
                                            " '" & sUsuario & "', " & _
                                                   chkResumen.Value & ", " & _
                                                   chkImpuesto1.Value & ", " & _
                                                   chkImpuesto2.Value & ", " & _
                                                   chkImpuesto3.Value & ", " & _
                                            " getdate()," & chkDocEquivDolares.Value & ", " & chLImprimeImageCab.Value & ", " & chLImprimeImagePie.Value & ", " & chkFacturacionE.Value & ",'" & txtAutorizacion.Text & "','" & Format(dtpFechaInicio.Value, "yyyy/MM/dd") & "','" & Format(dtpFechaCaducida.Value, "yyyy/MM/dd") & "','" & txtPrefijoEnlace.Text & "'," & chkFacturacionOfisis.Value & " ," & chkImpResumido.Value & "," & chkopGravInaf.Value & "," & Me.chkCodProdDes.Value & "," & Me.chkMayorCero.Value & "," & Me.chkSolicitarConfirmacion.Value & " )"
                       Else
                            Isql = "insert into TTIPODOCUMENTOIMPRESORA( " & _
                            "tCaja, tTipoEmision, tImpresora, tDescripcion, tFormulario, tSerie, tUltimoNumero, tUsuario, lResumen, lImpuesto1, lImpuesto2, lImpuesto3, fRegistro,lEquivaDolares,  lImprimeImageCab,lImprimeImagePie, lFacturacionElectronica,tPrefijoEnlace,lDocumentoElectronicoOfisis,tFormVenta,tCompVenta,limprimeresumen,lOpGravInaf, lImpProdDesc, limpDocMayorCero, lConfirmacionDoc) " & _
                            "values (  '" & txtCodigo.Text & "', " & _
                                     " '" & cboTipoDocumento.BoundText & "', " & _
                                     " '" & cboImpresora.BoundText & "', " & _
                                     " '" & txtDescripcion.Text & "', " & _
                                     " '" & cboFormulario.BoundText & "', " & _
                                     " '" & txtSerie.Text & "', " & _
                                     " '" & txtCorrelativo.Text & "', " & _
                                     " '" & sUsuario & "', " & _
                                            chkResumen.Value & ", " & _
                                            chkImpuesto1.Value & ", " & _
                                            chkImpuesto2.Value & ", " & _
                                            chkImpuesto3.Value & ", " & _
                                     " getdate()," & chkDocEquivDolares.Value & ", " & chLImprimeImageCab.Value & ", " & chLImprimeImagePie.Value & ", " & chkFacturacionE.Value & ",'" & txtPrefijoEnlace.Text & "'," & chkFacturacionOfisis.Value & ",'" & txtFormVenta.Text & "','" & txtCompVenta.Text & "'," & chkImpResumido.Value & "," & chkopGravInaf.Value & "," & Me.chkCodProdDes.Value & ", " & Me.chkMayorCero.Value & "," & Me.chkSolicitarConfirmacion.Value & " )"
                       End If
                       Cn.Execute Isql
                       RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsGrilla.Requery
                       RsGrilla.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
                       
               Else
                   sPasa = txtCodigo.Text
                                      
                   'Inserta Movimiento auditoria
                   If pais = "002" Then 'Ecuador
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "02", sUsuario, sPasa, "", _
                        "tcaja", "Codigo de Caja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                        "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                        "tSerie", "Serie", cboLocal.Text & txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                        "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.Value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), _
                        "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.Value, "Verdadero", "Falso"), _
                        "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.Value, "Verdadero", "Falso"), _
                        "lConfirmacionDoc", "Confirma Generacion de Documentos", IIf(chkSolicitarConfirmacion.Value, "Verdadero", "Falso"), _
                        "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.Value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.Value, "Verdadero", "Falso"))
                  Else
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "02", sUsuario, sPasa, "", _
                        "tcaja", "Codigo de Caja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                        "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                        "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                        "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.Value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), _
                        "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.Value, "Verdadero", "Falso"), _
                        "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.Value, "Verdadero", "Falso"), _
                        "lConfirmacionDoc", "Confirma Generacion de Documentos", IIf(chkSolicitarConfirmacion.Value, "Verdadero", "Falso"), _
                        "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.Value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.Value, "Verdadero", "Falso"))
                  End If
                  If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                   End If
               
                  'Cambiar el SQL
                  If pais = "002" Then 'Ecuador
                     Isql = "update TTIPODOCUMENTOIMPRESORA set " & _
                            "tImpresora ='" & cboImpresora.BoundText & "', " & _
                            "tDescripcion ='" & txtDescripcion.Text & "', " & _
                            "tFormulario ='" & cboFormulario.BoundText & "', " & _
                            "tSerie ='" & cboLocal.Text & txtSerie2.Text & "', " & _
                            "tUltimoNumero ='" & txtCorrelativo2.Text & "', " & _
                            "lResumen =" & chkResumen.Value & ", " & _
                            "lImpuesto1 =" & chkImpuesto1.Value & ", " & _
                            "lImpuesto2 =" & chkImpuesto2.Value & ", " & _
                            "lImpuesto3 =" & chkImpuesto3.Value & ", " & _
                            "lImprimeImageCab=" & chLImprimeImageCab.Value & ",  " & _
                            "lImprimeImagepie=" & chLImprimeImagePie.Value & ",  " & _
                            "lFacturacionElectronica=" & chkFacturacionE.Value & ",  " & _
                            "limprimeresumen=" & chkImpResumido.Value & ",  " & _
                            "lOpGravInaf=" & chkopGravInaf.Value & ",  " & _
                            "lDocumentoElectronicoOfisis =" & Me.chkFacturacionOfisis.Value & ", " & _
                            "lEquivaDolares =" & chkDocEquivDolares.Value & ", " & _
                            "fInicio = '" & Format(dtpFechaInicio.Value, "yyyy/MM/dd") & "', " & _
                            "tPrefijoEnlace ='" & txtPrefijoEnlace.Text & "', " & _
                            "fCaducidad = '" & Format(dtpFechaCaducida.Value, "yyyy/MM/dd") & "', " & _
                            "tNumeroAutorizacion ='" & txtAutorizacion.Text & "', " & _
                            "lImpProdDesc =" & Me.chkCodProdDes.Value & ", " & _
                            "lImpDocMayorCero =" & Me.chkMayorCero.Value & ", " & _
                            "lConfirmacionDoc =" & Me.chkSolicitarConfirmacion.Value & " " & _
                            " where tCaja = '" & txtCodigo.Text & "' and tTipoEmision = '" & cboTipoDocumento.BoundText & "'"
                   Else 'lImpProdDesc
                     Isql = "update TTIPODOCUMENTOIMPRESORA set " & _
                            "tImpresora ='" & cboImpresora.BoundText & "', " & _
                            "tDescripcion ='" & txtDescripcion.Text & "', " & _
                            "tFormulario ='" & cboFormulario.BoundText & "', " & _
                            "tSerie ='" & txtSerie.Text & "', " & _
                            "tUltimoNumero ='" & txtCorrelativo.Text & "', " & _
                            "lResumen =" & chkResumen.Value & ", " & _
                            "lImpuesto1 =" & chkImpuesto1.Value & ", " & _
                            "lImpuesto2 =" & chkImpuesto2.Value & ", " & _
                            "lImprimeImageCab=" & chLImprimeImageCab.Value & ",  " & _
                            "lImprimeImagepie=" & chLImprimeImagePie.Value & ",  " & _
                            "lImpuesto3 =" & chkImpuesto3.Value & ", " & _
                            "tPrefijoEnlace ='" & txtPrefijoEnlace.Text & "', " & _
                            "tFormVenta ='" & txtFormVenta.Text & "', " & _
                            "tCompVenta ='" & txtCompVenta.Text & "', " & _
                            "lFacturacionElectronica=" & chkFacturacionE.Value & ",  " & _
                            "limprimeresumen=" & chkImpResumido.Value & ",  " & _
                            "lOpGravInaf=" & chkopGravInaf.Value & ",  " & _
                            "lDocumentoElectronicoOfisis =" & Me.chkFacturacionOfisis.Value & ", " & _
                            "lEquivaDolares =" & chkDocEquivDolares.Value & ", " & _
                            "lImpProdDesc =" & Me.chkCodProdDes.Value & ", " & _
                            "lImpDocMayorCero =" & Me.chkMayorCero.Value & ", " & _
                            "lConfirmacionDoc =" & Me.chkSolicitarConfirmacion.Value & " " & _
                            " where tCaja = '" & txtCodigo.Text & "' and tTipoEmision = '" & cboTipoDocumento.BoundText & "'"
                   End If
                   Cn.Execute Isql
                   nPos = RsGrilla.AbsolutePosition 'chkImpResumido
                   RsGrilla.Requery
                   RsGrilla.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetalle True
          
          Case Is = 4 ' Cancelar
               SubDetalle True
               
          Case Is = 5 ' Agregar Area
               'Cambiar los Controles
                With RsAI
                     'Cuadro de Texto
                     cboArea.Text = ""
                     cboImpArea.Text = ""
                     txtDescripcion.Text = ""
                     lActivaDobleImpComanda.Value = 0
                     
                End With
                SubDetArea False
                wAgrega = True
          
          Case Is = 6 ' Modificar Area
               If RsAI.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetArea False
               wAgrega = False
               SubArea
          
          Case Is = 7 ' Eliminar Area
               If RsAI.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubArea
               If MsgBox("Seguro de Eliminar esta Area " & cboArea.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from TAREAIMPRESORA where tCaja ='" & txtCodigo.Text & "' and tArea ='" & cboArea.BoundText & "'"
               RsAI.Requery
               If RsAI.RecordCount <> 0 Then
                  RsAI.MoveLast
               End If
          
          Case Is = 8 ' Grabar ARea
               If wAgrega Then
                  If RsAI.RecordCount > 0 Then
                     RsAI.MoveFirst
                     RsAI.Find ("tArea ='" & cboArea.BoundText & "'")
                     If Not RsAI.EOF Then
                        MsgBox "Area " & cboArea.Text & " ya ingresada", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREAIMPRESORA( " & _
                          "tCaja, tArea, tImpresora, tUsuario, fRegistro,lActivaDobleImpComanda) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboArea.BoundText & "', " & _
                                   " '" & cboImpArea.BoundText & "', " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() " & ", " & _
                                   "" & lActivaDobleImpComanda.Value & ") "
            
                       Cn.Execute Isql
          
                       RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsAI.Requery
                       RsAI.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREAIMPRESORA set " & _
                         "tImpresora ='" & cboImpArea.BoundText & "',lActivaDobleImpComanda = " & lActivaDobleImpComanda.Value & _
                         " where tCaja = '" & txtCodigo.Text & "' and tArea = '" & cboArea.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = RsArea.AbsolutePosition
                   RsAI.Requery
                   RsAI.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetArea True
          
          Case Is = 9 ' Cancelar
               SubDetArea True
               
          'CESAR CHEF CONTROL
          'AREA CHEF
          Case Is = 10 ' Grabar AreaChef
               If wAgrega Then
                  If RsAChef.RecordCount > 0 Then
                     RsAChef.MoveFirst
                     RsAChef.Find ("Area ='" & cboAreaChef.BoundText & "'")
                     If Not RsAChef.EOF Then
                        MsgBox "Area " & cboAreaChef.Text & " ya ingresada", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREACHEF( " & _
                          "tCaja, tArea, lArea, tUsuario, fRegistro) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboAreaChef.BoundText & "', " & _
                                   " " & chkAreaChef.Value & " , " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
                       Cn.Execute Isql
          
                       RsAChef.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsAChef.Requery
                       RsAChef.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREACHEF set " & _
                         "lArea = " & chkAreaChef.Value & " " & _
                         " where tCaja = '" & txtCodigo.Text & "' and tArea = '" & cboAreaChef.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = RsArea.AbsolutePosition
                   RsAChef.Requery
                   RsAChef.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetAreaChef True
               
          Case Is = 11 'Cancelar Area Chef
               SubDetAreaChef True
               
          Case Is = 13 'Agregar Area Chef
                With RsAChef
                     cboAreaChef.Text = ""
                     chkAreaChef.Value = 0
                End With
                SubDetAreaChef False
                wAgrega = True
               
          Case Is = 14 ' Modificar Area Chef
               If RsAChef.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetAreaChef False
               wAgrega = False
               SubAreaChef
               
          Case Is = 12 ' Eliminar Area Chef
               If RsAChef.RecordCount = 0 Then
                  Exit Sub
               End If
               
               SubAreaChef
               If MsgBox("Seguro de Eliminar esta Area " & cboAreaChef.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from TAREACHEF where tCaja ='" & txtCodigo.Text & "' and tArea ='" & cboAreaChef.BoundText & "'"
               RsAChef.Requery
               If RsAChef.RecordCount <> 0 Then
                  RsAChef.MoveLast
               End If
          '---------------------------------------
               
          Case 15 ' graba sub grupos
                If cboSubGrupo.Text = "" Then
                        MsgBox "Seleccionar un Sub Grupo", vbInformation, sMensaje
                        Exit Sub
                End If
                If cboAreaProd.Text = "" Then
                        MsgBox "Seleccionar un Area", vbInformation, sMensaje
                        Exit Sub
                End If
                 If wAgrega Then
                  If rsAreaSubGrupo.RecordCount > 0 Then
                     rsAreaSubGrupo.MoveFirst
                     rsAreaSubGrupo.Find ("tSubGrupo ='" & cboSubGrupo.BoundText & "'")
                     If Not rsAreaSubGrupo.EOF Then
                        MsgBox "Sub Grupo " & cboSubGrupo.Text & " ya ingresado", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREASUBGRUPO( " & _
                          "TCAJA, TSUBGRUPO, TAREA,TUSUARIO,FREGISTRO) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboSubGrupo.BoundText & "', " & _
                                   " '" & cboAreaProd.BoundText & "', " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
                       Cn.Execute Isql
          
                       rsAreaSubGrupo.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       rsAreaSubGrupo.Requery
                       rsAreaSubGrupo.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREASUBGRUPO set " & _
                         "tarea ='" & cboAreaProd.BoundText & "' " & _
                         " where tCaja = '" & txtCodigo.Text & "' and tSubGrupo = '" & cboSubGrupo.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = rsAreaSubGrupo.AbsolutePosition
                   rsAreaSubGrupo.Requery
                   rsAreaSubGrupo.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetAreaSubGrupo True
            
           Case 16
               SubDetAreaSubGrupo True
        
         Case 17 ' lg eliminar:
          If rsAreaSubGrupo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubAreaSubGrupo
               If MsgBox("Seguro de Eliminar este Sub Grupo " & cboSubGrupo.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from tareasubgrupo where tCaja ='" & txtCodigo.Text & "' and tsubgrupo ='" & cboSubGrupo.BoundText & "'"
               rsAreaSubGrupo.Requery
               If rsAreaSubGrupo.RecordCount <> 0 Then
                  rsAreaSubGrupo.MoveLast
               End If
          Case 18 ' lg nuevp
              With rsAreaSubGrupo
                     cboSubGrupo.Text = ""
                     cboAreaProd.Text = ""
              End With
                SubDetAreaSubGrupo False
                wAgrega = True
        Case 19 ' lg modificacion area subgrupo
                If rsAreaSubGrupo.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetAreaSubGrupo False
               wAgrega = False
               SubAreaSubGrupo
   End Select
   
End Sub


Sub SubAreaSubGrupo()
    With rsAreaSubGrupo
         cboSubGrupo.BoundText = IIf(IsNull(!tSubGrupo), "", !tSubGrupo)
         cboAreaProd.BoundText = IIf(IsNull(!tArea), "", !tArea)
    End With
End Sub

Public Sub SubDetAreaSubGrupo(Activa As Boolean)
   fraAreaProduccion.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   'A-M-E
   cmdOpcionGrilla(18).Enabled = Activa
   cmdOpcionGrilla(19).Enabled = Activa
   cmdOpcionGrilla(17).Enabled = Activa
 
End Sub

Private Sub cmdQuitarFotoCabecera_Click()
    imgFoto.Picture = Nothing
    Cn.Execute "update tcaja set iimagencabdoc=null where tcaja='" & txtCodigo.Text & "'"
End Sub

Private Sub cmdQuitarFotoPie_Click()
    imgFotoPie.Picture = Nothing
    Cn.Execute "update tcaja set iimagenpiedoc=null where tcaja='" & txtCodigo.Text & "'"
End Sub

Private Sub Command1_Click()
Frame18.Visible = False
    
End Sub


Private Sub dgvBloqueoOrigenVenta_DblClick()
On Error GoTo fin:
    If MsgBox("Eliminar el bloqueo de origen de venta: " & RsBloqueoOrigenVenta!Descripcion & "?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
        Cn.Execute "delete from TCAJAORIGEN_BLOQUEO  where tcaja='" & Me.txtCodigo.Text & "' and CodigoOrigenventa='" & RsBloqueoOrigenVenta!codigo & "' "
         RsBloqueoOrigenVenta.Requery
    End If
fin:
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    
    'Ingrese el Titulo
    Me.Caption = " Mantenimiento de Cajas "
    fraDetalle.Caption = Me.Caption
    
    'Ingrese el SubTitulo
    grdGrilla.Caption = " Configuración de Documentos "
    fraGrilla.Visible = False
    
    grdAI.Caption = " Configuración de Areas "
    fraArea.Visible = False
    fraArea.Caption = grdAI.Caption
    
    If pais = "003" Then
        Label(25).Visible = True
        Label(25).Caption = "Impresora Fiscal: "
        cboComprobante.Visible = True
    Else
        Label(25).Visible = False
        cboComprobante.Visible = False
    End If
       
    'lucho areas por subgrupos
    Me.grdGrillaSubgrupos.Caption = " Configuración de Areas Por Sub Grupos "
    Me.fraAreaProduccion.Visible = False
    Me.fraAreaProduccion.Caption = grdGrillaSubgrupos.Caption
    
    
    'CESAR AREA CHEF
    grdAChef.Caption = " Configuración de Areas Chef Control "
    fraAreaChef.Visible = False
    fraAreaChef.Caption = grdAChef.Caption

    'Llena todos los Combos
    LlenaCombos
    
    'Ingresar la Vista de la Grilla
    Isql = "select * from vTipoDocumentoImpresora"
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(8, grdGrilla, "Descripción", 2, "Descripcion", 1900, 0, 0, "", _
                                  "Prefijo", 2, "Prefijo", 600, 2, 0, "", _
                                  "Serie", 2, "tSerie", 800, 2, 0, "", _
                                  "Número", 2, "tUltimoNumero", 1200, 0, 0, "", _
                                  "Impresora", 2, "Impresora", 1500, 0, 0, "", _
                                  "Autorizacion", 2, "tNumeroAutorizacion", 1400, 0, 0, "", _
                                  "Imag Cab", 2, "lImprimeImageCab", 920, 2, 4, "", _
                                  "Imag Pie", 2, "lImprimeImagePie", 920, 2, 4, "")
    Set grdGrilla.DataSource = RsGrilla
    
    'Ingresar la Vista de la Grilla 2
    Isql = "select * from vAreaImpresora"
    Set RsAI = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(2, grdAI, "Area", 2, "Area", 3000, 0, 0, "", _
                              "Impresora", 2, "Impresora", 3000, 0, 0, "")
    Set grdAI.DataSource = RsAI
    
    
    'lucho area subgrupo
        
    Isql = "Select * From vAreaSubGrupo"
    Set rsAreaSubGrupo = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdGrillaSubgrupos, "Caja", 2, "tCaja", 0, 0, 0, "", _
                                 "SubGrupo", 2, "SubGrupo", 3400, 0, 0, "", _
                                 "Area", 2, "Area", 3400, 0, 0, "")
                                 
    Set Me.grdGrillaSubgrupos.DataSource = rsAreaSubGrupo
    '------------------------------------
       
    'CESAR AREA CHEF
    Isql = "Select * From vAreaChef"
    Set RsAChef = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdAChef, "Caja", 2, "Caja", 2000, 0, 0, "", _
                                 "Area", 2, "Area", 3000, 0, 0, "", _
                                 "AreaChef", 2, "AreaChef", 1000, 2, 4, "")
    Set grdAChef.DataSource = RsAChef
    '------------------------------------
    
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       inicio
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
     'FE ARGENTINA GCAA 14/03/2022
   If pais = "003" Then
    Me.FrmFacArgentina.Visible = True
   End If
    
    
    If pais = "002" Then 'Ecuador
       txtSerie.Visible = False
       txtCorrelativo.Visible = False
       cboLocal.Visible = True
       txtSerie2.Visible = True
       txtCorrelativo2.Visible = True
              
       If lFacturacionE Then
            txtAutorizacion.Enabled = False
            dtpFechaInicio.Enabled = False
            dtpFechaCaducida.Enabled = False
       Else
            txtAutorizacion.Enabled = True
            dtpFechaInicio.Enabled = True
            dtpFechaCaducida.Enabled = True
       End If
   
    Else
       txtSerie.Visible = True
       txtCorrelativo.Visible = True
       cboLocal.Visible = True
       txtSerie2.Visible = False
       txtCorrelativo2.Visible = False
       cboLocal.Visible = False
       Label(19).Visible = False
       Label(26).Visible = False
       Label(27).Visible = False
       
       txtAutorizacion.Visible = False
       dtpFechaInicio.Visible = False
       dtpFechaCaducida.Visible = False
       
    End If
    
    'Configuración de la Grilla
    Call ConfGrilla(3, dgvBloqueoOrigenVenta, "CanalVenta", 2, "CanalVenta", 1300, 2, 0, "", _
                                    "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                  "Descripcion", 2, "Descripcion", 2000, 0, 0, "")
    
    
    Me.tabOpcion.Tab = 0
    cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsGrupo = Nothing
   Set RsImpresora = Nothing
   Set RsPreCuenta = Nothing
   Set RsTipoDocumento = Nothing
   Set RsArea = Nothing
   Set RsAI = Nothing
   Set RsImpArea = Nothing
   Set RsFormulario = Nothing
   Set RsGrilla = Nothing
   Set RsTipoPedido = Nothing
   Set frmCajaDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
    'btnNewModal.Enabled = Activa
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcionGrilla_Click (1)
End Sub

Public Sub SubDetalle(Activa As Boolean)
   With cboImpresora
        Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
        Set RsImpresora = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsImpresora
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
   End With

   fraGrilla.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   cmdOpcionGrilla(0).Enabled = Activa
   cmdOpcionGrilla(1).Enabled = Activa
   cmdOpcionGrilla(2).Enabled = Activa

   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub






















Private Sub txtBalanzaPuerto_Change()
    If Not IsNumeric(Me.txtBalanzaPuerto.Text) And Trim(Me.txtBalanzaPuerto.Text) <> "" Then
        MsgBox "Valor no valido!!!", vbInformation, sMensaje
        Me.txtBalanzaPuerto.Text = ""
        Me.txtBalanzaPuerto.SetFocus
    End If
End Sub

Private Sub txtbaltiempo_Change()
    If Not IsNumeric(Me.txtbaltiempo.Text) And Trim(Me.txtbaltiempo.Text) <> "" Then
        MsgBox "Valor no valido!!!", vbInformation, sMensaje
        Me.txtbaltiempo.Text = ""
        Me.txtbaltiempo.SetFocus
    End If
End Sub

Private Sub txtbaltiempo_LostFocus()
'    If Not IsNumeric(Me.txtbaltiempo.Text) And Trim(Me.txtbaltiempo.Text) <> "" Then
'        MsgBox "Valor no valido!!!", vbInformation, sMensaje
'    End If
End Sub

Private Sub txtCorrelativo_Lostfocus()
   txtCorrelativo.Text = Mid("000000000", 1, 9 - Len(Trim(str(val(txtCorrelativo.Text))))) + Trim(str(val(txtCorrelativo.Text)))
End Sub

Private Sub txtSerie_LostFocus()
   'txtSerie.Text = Mid("00000", 1, 5 - Len(Trim(str(Val(txtSerie.Text))))) + Trim(str(Val(txtSerie.Text)))
   txtSerie.Text = UCase(txtSerie.Text)
End Sub

Private Sub txtCorrelativo2_Lostfocus()
   txtCorrelativo2.Text = Mid("00000000", 1, 9 - Len(Trim(str(val(txtCorrelativo2.Text))))) + Trim(str(val(txtCorrelativo2.Text)))
End Sub

Private Sub txtSerie2_LostFocus()
   txtSerie2.Text = Mid("000", 1, 3 - Len(Trim(str(val(txtSerie2.Text))))) + Trim(str(val(txtSerie2.Text)))
End Sub

Sub SubAsignar()
    With RsGrilla
         'Cuadro de Texto
         cboTipoDocumento.BoundText = IIf(IsNull(!TTipoEmision), "", !TTipoEmision)
         cboImpresora.BoundText = IIf(IsNull(!timpresora), "", Trim(!timpresora))
         cboFormulario.BoundText = IIf(IsNull(!tFormulario), "", Trim(!tFormulario))
         txtDescripcion.Text = IIf(IsNull(!tDescripcion), "", !tDescripcion)
         chkResumen.Value = IIf(IsNull(!lResumen), 0, IIf(!lResumen, 1, 0))
         chLImprimeImageCab.Value = IIf(IsNull(!lImprimeImageCab), 0, IIf(!lImprimeImageCab, 1, 0))
         chLImprimeImagePie.Value = IIf(IsNull(!lImprimeImagepie), 0, IIf(!lImprimeImagepie, 1, 0))
         chkImpResumido.Value = IIf(IsNull(!lImprimeResumen), 0, IIf(!lImprimeResumen, 1, 0))
         chkSolicitarConfirmacion.Value = IIf(IsNull(!lConfirmacionDoc), 0, IIf(!lConfirmacionDoc, 1, 0))
         
        
         'FACTURACION ELECTRONICA
         chkFacturacionE.Value = IIf(IsNull(!lFacturacionElectronica), 0, IIf(!lFacturacionElectronica, 1, 0))
         
         chkFacturacionOfisis.Value = IIf(IsNull(!lDocumentoElectronicoOfisis), 0, IIf(!lDocumentoElectronicoOfisis, 1, 0))
         chkopGravInaf.Value = IIf(IsNull(!lOpGravInaf), 0, IIf(!lOpGravInaf, 1, 0))
         Me.chkCodProdDes.Value = IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))  'IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))
         Me.chkMayorCero.Value = IIf(IsNull(!lImpDocMayorCero), 0, IIf(!lImpDocMayorCero, 1, 0))  'IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))
         If pais = "002" Then 'Ecuador
            cboLocal.Text = Mid(IIf(IsNull(!tSerie), "001", !tSerie), 1, 3)
            txtSerie2.Text = Mid(IIf(IsNull(!tSerie), "001", !tSerie), 4, 3)
            txtCorrelativo2.Text = IIf(IsNull(!tUltimoNumero), "", !tUltimoNumero)
            txtAutorizacion.Text = IIf(IsNull(!tNumeroAutorizacion), "", !tNumeroAutorizacion)
            dtpFechaInicio.Value = IIf(IsNull(!fInicio), FechaServidor(), !fInicio)
            dtpFechaCaducida.Value = IIf(IsNull(!fCaducidad), FechaServidor(), !fCaducidad)
         Else
            txtSerie.Text = IIf(IsNull(!tSerie), "", !tSerie)
            txtCorrelativo.Text = IIf(IsNull(!tUltimoNumero), "", !tUltimoNumero)
            txtAutorizacion.Text = ""
         End If
         
         txtPrefijoEnlace.Text = IIf(IsNull(!tPrefijoEnlace), "", !tPrefijoEnlace)
         
         txtFormVenta.Text = IIf(IsNull(!tFormVenta), "", !tFormVenta)
         txtCompVenta.Text = IIf(IsNull(!tCompVenta), "", !tCompVenta)
         
         If IsNull(!LEQUIVADOLARES) = True Then
            chkDocEquivDolares.Value = 0
         ElseIf !LEQUIVADOLARES = 0 Then
            chkDocEquivDolares.Value = 0
        Else
            chkDocEquivDolares.Value = 1
         End If
                  
         If !TTipoEmision = "00" Then
            Frame14.Visible = False
            'Label(4).Visible = False
         Else
            'Label(4).Visible = true
            Frame14.Visible = True
         End If
        
         If sImpuesto1 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto1.Visible = True
            chkImpuesto1.Caption = sImpuesto1
            chkImpuesto1.Value = IIf(IsNull(!lImpuesto1), 0, IIf(!lImpuesto1, 1, 0))
         Else
            chkImpuesto1.Visible = False
            chkImpuesto1.Value = 0
         End If
             
         If sImpuesto2 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto2.Visible = True
            chkImpuesto2.Caption = sImpuesto2
            chkImpuesto2.Value = IIf(IsNull(!lImpuesto2), 0, IIf(!lImpuesto2, 1, 0))
         Else
            chkImpuesto2.Visible = False
            chkImpuesto2.Value = 0
         End If
             
         If sImpuesto3 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto3.Visible = True
            chkImpuesto3.Caption = sImpuesto3
            chkImpuesto3.Value = IIf(IsNull(!lImpuesto3), 0, IIf(!lImpuesto3, 1, 0))
        Else
           chkImpuesto3.Visible = False
           chkImpuesto3.Value = 0
        End If
    End With
End Sub

Public Sub SubDetArea(Activa As Boolean)
   With cboImpArea
        Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
        Set RsImpArea = Lib.OpenRecordset(Isql, Cn)
        
         Set .RowSource = RsImpArea
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
   End With

   fraArea.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   
   cmdOpcionGrilla(5).Enabled = Activa
   cmdOpcionGrilla(6).Enabled = Activa
   cmdOpcionGrilla(7).Enabled = Activa
   
   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub

Sub SubArea()
    With RsAI
         cboArea.BoundText = IIf(IsNull(!tArea), "", !tArea)
         cboImpArea.BoundText = IIf(IsNull(!timpresora), "", Trim(!timpresora))
         lActivaDobleImpComanda.Value = IIf(IsNull(!lActivaDobleImpComanda), 0, IIf(!lActivaDobleImpComanda = True, 1, 0))
    End With
End Sub

Public Sub inicio()
    RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
    RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
    
      
    rsAreaSubGrupo.Filter = "tcaja='" & txtCodigo.Text & "'"
    
    cmdOpcionGrilla(0).Enabled = False
    cmdOpcionGrilla(1).Enabled = False
    cmdOpcionGrilla(2).Enabled = False
    cmdOpcion(0).Enabled = False
    cmdOpcion(2).Enabled = False
    
    chkVComanda.Value = 0
    chkComanda.Value = 0
    chkObligaPrinter.Value = 0
    chkObligaPrecuenta.Value = 0
   ' chkObliga.value = 0
   'chkMozo.value = 0
  '  chkMotorizado.value = 0
    chkAdicion.Value = 0
   ' chkPax.value = 0
    chkConsumo1.Value = 0
    chkConsumo2.Value = 0
    chkConsumo3.Value = 0
    
    chkCodigoReciboIngreso.Value = 0
    
    chkComboPrecuenta.Value = 0
    chkEliminaC.Value = 0
    chkPasswordC.Value = 0
    chkElimina.Value = 0
    chkPassword.Value = 0
    chkAccesoDespachoPedido.Value = 0
    chkObligaCierre.Value = 0
    chkFiltroTipoPedido.Value = 0
    chkEquivaPrecuenta = 0
    chkCancelacion.Value = 0
    chkDirecto.Value = 0
    chkCambioMesa.Value = 0
    chkPrecuenta.Value = 0
    chkBuscaPedido.Value = 0
    chkImprimeImagCabPrecuenta.Value = 0
    chkImprimeImagPiePrecuenta.Value = 0
    chkAgrupada.Value = 0
    chkSiab.Value = 0
    chkComboDocumento.Value = 0
    chkVisaNet.Value = 0
    chkImpuestoPrecuenta.Value = 0
    chkDocumentoAgrupado.Value = 0
    chkOrden.Value = 0
    chkActivo.Value = 1
    chkValor.Value = 0
    chkCajaMobile.Value = 0
    txtLimitePrecuenta.Text = "0"
    txtLimiteReimpresion.Text = "0"
    chkPasswordTransferencia.Value = 0
    chkPasswordImportar.Value = 0
    chkDescripcionAlternativa.Value = 0
    Me.chkMotDesc.Value = 0
    Me.chkCajaContingencia.Value = 0
    Me.chkImpPropina.Value = 0
    chkCompatibilidadTVS.Value = 0 'TVS
    chkCD.Value = 0
    chkDisgrega.Value = 0
    chkMultiCajero.Value = 0
    chkMCPV.Value = 0
   ' chkFechaDelivery.value = 0
    chkCCVOX.Value = 0
    chkObservacionPrecuenta.Value = 0
    chkObservacionDocumento.Value = 0
    chkObservacionCabDoc.Value = 0
    txtLongitudBarra.Text = "0"
    chkPagoRapido.Value = 0
    
    chkPasswordPorCobrar.Value = 0
    chkModificaTipoPedido.Value = 0
    txtBalanzaPuerto.Text = ""
    chkBloqueaPrecuenta.Value = 0
    txtLimitePrecuenta.Enabled = True
    chkMulti1.Value = 0
    chkMulti2.Value = 0
    Me.chkBuscaPedidoFiltrarMesa.Value = 0
    Me.chkBuscaPedidoVisualizaGrilla.Value = 0
    Me.chkPassOtrosPagos.Value = 0
    Me.chkImpClienteCab.Value = 0
    Me.chkBuscarPedidoBarra.Value = 0
    Me.ChkActivaCashDro.Value = 0
    Me.chkActivaUber.Value = 0
    Me.chkIntMercadoPagoQR.Value = 0
    Me.chkIntMercadoPagoQREstatico.Value = 0
    Me.chkIntVisaNetQR.Value = 0
    Me.chkIntegracionPatio.Value = 0
    Me.txtIpCashDrow.Text = "192.168.1.1"
    Me.chkPagoEfectivo.Value = 0
    Me.ChkPagoDirectoCD.Value = 0
    Me.chkImprimePropinaSugerida.Value = 0
    Me.chklActivarAvisoRetenciones.Value = 0
    Me.chklActivaTresDecimales.Value = 0
    Me.chklIntegracionCASSWII30CD.Value = 0
    Me.chklActivarMensajeCtsCorrientes.Value = 0
    Me.ChlNoVisualizarItemCombos.Value = 0
    Me.chklActivaTresDecimales.Value = 0
    Me.ChlMostrarPrecioItemsCombo.Value = 0
    Me.chklIntMobileKDS.Value = 0
    Me.chkGuiaF2.Value = 0
    
    Me.chklPermitirAnularMaxDias.Value = 0
    Me.chklOrdenarClientes.Value = 0
    
    Me.chkIntegracionWebApi.Value = 0
    
    Me.ChklActivaRestriClienteDelivery.Value = 0
    
End Sub

'lg

Private Sub chkMulti1_Click()
    If chkMulti1.Value Then
        chkMulti2.Value = 0
        fra2.Enabled = False
        fra1.Enabled = True
    Else
            fra1.Enabled = False
    End If
End Sub

Private Sub chkMulti2_Click()
    If chkMulti2.Value Then
        chkMulti1.Value = 0
        fra1.Enabled = False
        fra2.Enabled = True
    Else
        fra2.Enabled = False
    End If
End Sub

'CESAR AREA CHEF
Public Sub SubDetAreaChef(Activa As Boolean)
   fraAreaChef.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   'A-M-E
   cmdOpcionGrilla(13).Enabled = Activa
   cmdOpcionGrilla(14).Enabled = Activa
   cmdOpcionGrilla(12).Enabled = Activa
   
   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub
Sub SubAreaChef()
    With RsAChef
         cboAreaChef.BoundText = IIf(IsNull(!tArea), "", !tArea)
         chkAreaChef = IIf(IsNull(!AreaChef), 0, IIf(!AreaChef, 1, 0))
    End With
End Sub
'------------------------
