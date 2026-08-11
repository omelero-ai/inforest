VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   ForeColor       =   &H00808080&
   Icon            =   "frmParametro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnKDS 
      Caption         =   "Ver Mas"
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
      Left            =   60
      TabIndex        =   148
      Top             =   8775
      Width           =   1305
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8685
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   15319
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmParametro.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ChkCambiarContrasena"
      Tab(0).Control(1)=   "fraRuc"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmParametro.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame16"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame(8)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame(17)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame(16)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame(9)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame(15)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame(13)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame(10)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Frame(14)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Frame(11)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame(12)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame(19)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Complementos"
      TabPicture(2)   =   "frmParametro.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "ChklActivaMotivoAnulacion"
      Tab(2).Control(3)=   "Frame40"
      Tab(2).Control(4)=   "Frame39"
      Tab(2).Control(5)=   "Frame24"
      Tab(2).Control(6)=   "FrmVisor"
      Tab(2).Control(7)=   "Frame14"
      Tab(2).Control(8)=   "Frame13"
      Tab(2).Control(9)=   "frmMobile"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Facturacion Electronica"
      TabPicture(3)   =   "frmParametro.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmFacArgentina"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Leyes Aplicables"
      TabPicture(4)   =   "frmParametro.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame51"
      Tab(4).Control(1)=   "Frame45"
      Tab(4).Control(2)=   "Frame44"
      Tab(4).Control(3)=   "Frame37"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Integraciones"
      TabPicture(5)   =   "frmParametro.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame38"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Tiempo Limite Sesion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   -69240
         TabIndex        =   366
         Top             =   7560
         Width           =   4575
         Begin VB.TextBox txtTimer 
            Height          =   285
            Left            =   2160
            TabIndex        =   369
            Top             =   550
            Width           =   2055
         End
         Begin VB.CheckBox chkTimer 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa timer expira sesion"
            Height          =   255
            Left            =   120
            TabIndex        =   367
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Tiempo maximo de espera :"
            Height          =   255
            Left            =   120
            TabIndex        =   368
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Inforest App"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   -69240
         TabIndex        =   363
         Top             =   6435
         Width           =   4575
         Begin VB.TextBox tUrlImgInforestApp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   364
            Top             =   540
            Width           =   4200
         End
         Begin VB.Label Label 
            Caption         =   "Imagen (URL) :"
            Height          =   255
            Index           =   95
            Left            =   120
            TabIndex        =   365
            Top             =   285
            Width           =   2040
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Recibo de ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   6840
         TabIndex        =   361
         Top             =   7560
         Width           =   3495
         Begin VB.CheckBox chkReciboSaldoFavor 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa impresion de QR y uso de saldo a favor en pago de documentos."
            Height          =   495
            Left            =   120
            TabIndex        =   362
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.CheckBox ChkCambiarContrasena 
         Caption         =   "Cambiar Contraseña"
         Height          =   255
         Left            =   -66745
         TabIndex        =   353
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Frame Frame51 
         Caption         =   "Detracción"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   342
         Top             =   3120
         Width           =   5415
         Begin VB.CheckBox ChklDetraccion 
            Caption         =   "Aplicar Detracción"
            Height          =   255
            Left            =   120
            TabIndex        =   346
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox TxttMsjDetraccion 
            Height          =   1095
            Left            =   120
            MaxLength       =   200
            TabIndex        =   345
            Top             =   1200
            Width           =   5175
         End
         Begin VB.TextBox TxtlMontoDetraccion 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4200
            TabIndex        =   344
            Text            =   "750"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtlPorcentajeDetraccion 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4200
            TabIndex        =   343
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Nota: Recordar que debe estar activo el check de registro de pago antes de impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   120
            TabIndex        =   349
            Top             =   2400
            Width           =   5055
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje detracción (%):"
            Height          =   195
            Index           =   93
            Left            =   2300
            TabIndex        =   348
            Top             =   810
            Width           =   1860
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Facturas mayor igual a:"
            Height          =   195
            Index           =   91
            Left            =   2520
            TabIndex        =   347
            Top             =   320
            Width           =   1635
         End
      End
      Begin VB.CheckBox ChklActivaMotivoAnulacion 
         Caption         =   "Activar Motivo de Anulación al anular un documento"
         Height          =   315
         Left            =   -69240
         TabIndex        =   339
         Top             =   6000
         Width           =   4575
      End
      Begin VB.Frame FrmFacArgentina 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   335
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame Frame48 
            Caption         =   "Integracion TUSFACTURASAPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   720
            Left            =   120
            TabIndex        =   336
            Top             =   240
            Width           =   4820
            Begin VB.CheckBox chkFETusFacturasAPP 
               Caption         =   "Facturación Electrónica TusFacturasAPP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   337
               Top             =   240
               Width           =   4620
            End
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Imprimir Foma Pago FAC / BOL Electronica"
         Height          =   615
         Index           =   12
         Left            =   6840
         TabIndex        =   331
         Top             =   6960
         Width           =   3495
         Begin VB.CheckBox ChklImprimirFormaPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Imprimir Forma de Pago:"
            Height          =   255
            Left            =   120
            TabIndex        =   332
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   11
         Left            =   6840
         TabIndex        =   324
         Top             =   6360
         Width           =   3525
         Begin VB.CheckBox ChlReservaPedido 
            Alignment       =   1  'Right Justify
            Caption         =   "Convertir Reserva a Pedido"
            Height          =   255
            Left            =   120
            TabIndex        =   325
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame45 
         Caption         =   "Monto Minimo para la Retención:"
         Height          =   735
         Left            =   -72000
         TabIndex        =   322
         Top             =   2280
         Width           =   2535
         Begin VB.TextBox TxtMontoMinRete 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   240
            MaxLength       =   3
            TabIndex        =   323
            Text            =   "700"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame44 
         Caption         =   "Ingresar % Porcentaje de Retención"
         Height          =   735
         Left            =   -74880
         TabIndex        =   320
         Top             =   2280
         Width           =   2775
         Begin VB.TextBox TxtlRetencion 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            MaxLength       =   2
            TabIndex        =   321
            Text            =   "0"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Precuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   14
         Left            =   3600
         TabIndex        =   317
         Top             =   7515
         Width           =   3135
         Begin VB.TextBox txtPorcentajePropinaSugerida 
            Alignment       =   1  'Right Justify
            Height          =   290
            Left            =   2160
            TabIndex        =   319
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Porcentaje de propina sugerida en precuenta (%):"
            Height          =   495
            Left            =   120
            TabIndex        =   318
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Glosa Exonerado e Inafecto"
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
         Index           =   10
         Left            =   6840
         TabIndex        =   309
         Top             =   4680
         Width           =   3525
         Begin VB.TextBox txttGlosaImprExonerado 
            Height          =   735
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   314
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txttGlosaImprInafecto 
            Height          =   735
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   311
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label 
            Caption         =   "Glosa Impresion Exonerado:"
            Height          =   735
            Index           =   86
            Left            =   90
            TabIndex        =   313
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Glosa Impresion Inafecto:"
            Height          =   735
            Index           =   87
            Left            =   90
            TabIndex        =   310
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame40 
         Caption         =   "Api Web"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1575
         Left            =   -74880
         TabIndex        =   299
         Top             =   5880
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtUrlApi 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   301
            Top             =   600
            Width           =   5295
         End
         Begin VB.TextBox txtCodigoTiendaApi 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   300
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label 
            Caption         =   "Ruta: Ejemplo(http://192.168.1.1:80/WebApiInfomatica/api)"
            Height          =   255
            Index           =   84
            Left            =   120
            TabIndex        =   303
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label Label 
            Caption         =   "Codigo asignado en Api Web"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   302
            Top             =   960
            Width           =   4215
         End
      End
      Begin VB.Frame Frame39 
         Caption         =   "Imagenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1740
         Left            =   -69240
         TabIndex        =   279
         Top             =   4200
         Width           =   4575
         Begin VB.TextBox TxtRutaImaDia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   283
            Top             =   1200
            Width           =   4215
         End
         Begin VB.TextBox TxtRutaLogo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   282
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label 
            Caption         =   "Imagen del Día:"
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   281
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label 
            Caption         =   "Logo:"
            Height          =   255
            Index           =   82
            Left            =   120
            TabIndex        =   280
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7575
         Left            =   -74880
         TabIndex        =   274
         Top             =   360
         Width           =   10095
         Begin VB.Frame Frame47 
            Caption         =   "Bermur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1095
            Left            =   5760
            TabIndex        =   326
            Top             =   2640
            Width           =   4215
            Begin VB.TextBox TxtMontoMaximo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   333
               Text            =   "20"
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtCodigoTarjeta 
               Height          =   285
               Left            =   1800
               TabIndex        =   328
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton cmdBuscaTarjeta 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3360
               Picture         =   "frmParametro.frx":00B4
               Style           =   1  'Graphical
               TabIndex        =   327
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label 
               Caption         =   "Monto maximo de Recarga:"
               Height          =   435
               Index           =   78
               Left            =   240
               TabIndex        =   334
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Cod de Item Tarjeta:"
               Height          =   195
               Index           =   71
               Left            =   240
               TabIndex        =   329
               Top             =   240
               Width           =   1440
            End
         End
         Begin VB.Frame Frame41 
            Caption         =   "Patio de comidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1380
            Left            =   5760
            TabIndex        =   304
            Top             =   1080
            Width           =   4245
            Begin VB.CommandButton cmdItemDelivery 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3480
               Picture         =   "frmParametro.frx":01B6
               Style           =   1  'Graphical
               TabIndex        =   307
               Top             =   600
               Width           =   630
            End
            Begin VB.CheckBox ChkItemDelivery 
               Alignment       =   1  'Right Justify
               Caption         =   "Activar ingreso de costo de delivery:"
               Height          =   255
               Left            =   120
               TabIndex        =   306
               Top             =   240
               Width           =   3975
            End
            Begin VB.TextBox txtItemDelivery 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1920
               TabIndex        =   305
               Tag             =   "02155454555"
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label 
               Caption         =   "Item de costo delivery:"
               Height          =   255
               Index           =   70
               Left            =   120
               TabIndex        =   308
               Top             =   600
               Width           =   1695
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Enlace Mesa247"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1140
            Left            =   120
            TabIndex        =   294
            Top             =   2640
            Visible         =   0   'False
            Width           =   5505
            Begin VB.TextBox txtCajaMesa247 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2520
               MaxLength       =   3
               TabIndex        =   296
               Top             =   360
               Width           =   1320
            End
            Begin VB.TextBox txtAdicionMesa247 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2520
               MaxLength       =   3
               TabIndex        =   295
               Top             =   705
               Width           =   1320
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Caja Recepción de Pedido :"
               Height          =   195
               Index           =   57
               Left            =   240
               TabIndex        =   298
               Top             =   405
               Width           =   2130
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Adición Recepción de Pedido :"
               Height          =   255
               Index           =   60
               Left            =   120
               TabIndex        =   297
               Top             =   720
               Width           =   2250
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Enlace SAP "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1425
            Left            =   120
            TabIndex        =   284
            Top             =   1080
            Width           =   5535
            Begin VB.CheckBox ChkSAP 
               Caption         =   "Integración SAP"
               Height          =   375
               Left            =   240
               TabIndex        =   293
               Top             =   240
               Width           =   2295
            End
            Begin VB.Frame Frame19 
               Caption         =   "Datos del Servidor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   120
               TabIndex        =   287
               Top             =   2910
               Width           =   5295
               Begin VB.TextBox txtServidorSAP 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   290
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox TxtBaseSAP 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   289
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.CommandButton cmdConSAP 
                  Caption         =   "Probar Conexion"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   4230
                  Style           =   1  'Graphical
                  TabIndex        =   288
                  Top             =   330
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Servidor :"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   292
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label2 
                  Caption         =   "Base de Datos :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   291
                  Top             =   720
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame20 
               Caption         =   "Datos Local - Ingresar Codigo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   285
               Top             =   600
               Width           =   3375
               Begin VB.TextBox TxtCodAlmcSAP 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   120
                  TabIndex        =   286
                  Top             =   240
                  Width           =   2895
               End
            End
         End
         Begin VB.TextBox txtRutaImgFE 
            Height          =   285
            Left            =   2520
            TabIndex        =   278
            Top             =   600
            Width           =   7515
         End
         Begin VB.TextBox txtRutaMotorIntegraciones 
            Height          =   285
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   277
            Top             =   240
            Width           =   7515
         End
         Begin VB.Label Label 
            BackColor       =   &H008080FF&
            Caption         =   "Ruta de archivos o image:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   120
            TabIndex        =   276
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label 
            BackColor       =   &H008080FF&
            Caption         =   "Ruta Motor Integraciones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   275
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame37 
         Caption         =   "Leyes             Descripcion Resumido      Porcentaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   -74880
         TabIndex        =   260
         Top             =   480
         Width           =   10215
         Begin VB.TextBox txtLey1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   266
            ToolTipText     =   "Impuesto al consumo."
            Top             =   285
            Width           =   2310
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   265
            Top             =   720
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   264
            Top             =   1155
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.TextBox txtValorLey1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3570
            TabIndex        =   263
            ToolTipText     =   "Impuesto Aplicado a cada unidad en el comprobante de venta sin afectar el neto de venta."
            Top             =   285
            Width           =   690
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3570
            TabIndex        =   262
            Top             =   720
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3570
            TabIndex        =   261
            Top             =   1155
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label11 
            Caption         =   "(Impuesto al consumo por unidad y monto sin afectar el neto de venta)"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4680
            TabIndex        =   273
            Top             =   315
            Width           =   5055
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Ley 1 :"
            Height          =   195
            Index           =   74
            Left            =   330
            TabIndex        =   272
            Top             =   285
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   77
            Left            =   4320
            TabIndex        =   271
            Top             =   300
            Width           =   300
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   76
            Left            =   4320
            TabIndex        =   270
            Top             =   735
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   75
            Left            =   4320
            TabIndex        =   269
            Top             =   1170
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Ley 2 :"
            Height          =   195
            Index           =   73
            Left            =   330
            TabIndex        =   268
            Top             =   735
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Ley 3 :"
            Height          =   195
            Index           =   72
            Left            =   330
            TabIndex        =   267
            Top             =   1200
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   13
         Left            =   3600
         TabIndex        =   225
         Top             =   6480
         Width           =   3165
         Begin VB.CheckBox chkBloqInafecto 
            Alignment       =   1  'Right Justify
            Caption         =   "No permitir comandar Platos Inafectos y afectos en un mismo pedido."
            Height          =   615
            Left            =   120
            TabIndex        =   226
            ToolTipText     =   "Si esta Activo el check no se permitira comandar Items Inafectos en pedido Afecto."
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Descargo de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Index           =   15
         Left            =   3600
         TabIndex        =   221
         Top             =   4500
         Width           =   3165
         Begin VB.CheckBox chkControlInventario 
            Alignment       =   1  'Right Justify
            Caption         =   "Validar cierres de inventarios"
            Height          =   375
            Left            =   120
            TabIndex        =   360
            Top             =   960
            Width           =   2805
         End
         Begin VB.CheckBox chkDescargoInsumo 
            Alignment       =   1  'Right Justify
            Caption         =   "Descargo de Ventas por Insumo"
            Height          =   375
            Left            =   120
            TabIndex        =   253
            Top             =   600
            Width           =   2805
         End
         Begin VB.CheckBox chkValidaStock 
            Alignment       =   1  'Right Justify
            Caption         =   "Validar Stock de Insumos en Descargo"
            Height          =   375
            Left            =   120
            TabIndex        =   222
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Anticipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Index           =   9
         Left            =   6840
         TabIndex        =   215
         Top             =   3700
         Width           =   3525
         Begin VB.CommandButton cmdBuscaAnticipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            Picture         =   "frmParametro.frx":02B8
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   600
            Width           =   630
         End
         Begin VB.CheckBox chkActivaAnticipo 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar Anticipo"
            Height          =   255
            Left            =   120
            TabIndex        =   218
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtCodigoAnticipo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   217
            Tag             =   "02155454555"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label 
            Caption         =   "Cod de Item Anticipo:"
            Height          =   255
            Index           =   88
            Left            =   120
            TabIndex        =   216
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Caja Rapida - Pago de Documentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   16
         Left            =   120
         TabIndex        =   205
         Top             =   7380
         Width           =   3360
         Begin VB.CheckBox chkPagoCheque 
            Alignment       =   1  'Right Justify
            Caption         =   "Desactivar Forma de pago Cheque/Depositos :"
            Height          =   435
            Left            =   120
            TabIndex        =   207
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   3090
         End
         Begin VB.CheckBox chkPagoOtra 
            Alignment       =   1  'Right Justify
            Caption         =   "Desactivar Otras Formas de Pago :"
            Height          =   195
            Left            =   120
            TabIndex        =   206
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   680
            Width           =   3090
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Facturación Electrónica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7515
         Left            =   -74850
         TabIndex        =   185
         Top             =   480
         Width           =   10185
         Begin VB.CheckBox chkNotaCero 
            Caption         =   "Habilitar Notas de Crédito en monto 0"
            Height          =   255
            Left            =   360
            TabIndex        =   372
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CheckBox chkSolicitarDato 
            Caption         =   "Boleta con o sin datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   352
            Top             =   1920
            Width           =   4455
         End
         Begin VB.TextBox txtNumdiasAnulacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3570
            TabIndex        =   350
            Text            =   "0"
            Top             =   1400
            Width           =   735
         End
         Begin VB.Frame FrmFacEcuador 
            Height          =   7215
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   5055
            Begin VB.Frame FrmCamposFacturacion 
               Caption         =   "Cabecera Facturacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   975
               Left            =   120
               TabIndex        =   370
               Top             =   2520
               Width           =   4815
               Begin VB.CheckBox chkCabeceraFacturacion 
                  Caption         =   "Activar visualizacion de campos autorizacion documentos"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   371
                  Top             =   480
                  Width           =   4455
               End
            End
            Begin VB.Frame Frame36 
               Caption         =   "Integracion Estupendo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   600
               Left            =   120
               TabIndex        =   250
               Top             =   240
               Width           =   4820
               Begin VB.CheckBox chkFEEstupendo 
                  Caption         =   "Facturación Electrónica Estupendo"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   251
                  Top             =   300
                  Width           =   3060
               End
            End
         End
         Begin VB.Frame FrmFacPeru 
            Height          =   7335
            Left            =   5040
            TabIndex        =   227
            Top             =   120
            Visible         =   0   'False
            Width           =   5055
            Begin VB.Frame FrameOfiisis 
               Caption         =   "Integración Ofisis"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   615
               Left            =   120
               TabIndex        =   247
               Top             =   120
               Width           =   4820
               Begin VB.CheckBox chkFEOfisis 
                  Caption         =   "Facturación Electrónica Ofisis"
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  Width           =   2535
               End
               Begin VB.Label Label 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion a base de datos Sql Server"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   61
                  Left            =   1710
                  TabIndex        =   249
                  Top             =   150
                  Width           =   2985
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame Frame26 
               Caption         =   "Integracion Spring"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   610
               Left            =   120
               TabIndex        =   244
               Top             =   700
               Width           =   4820
               Begin VB.CheckBox chkFESpring 
                  Caption         =   "Facturación Electrónica Spring"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   245
                  Top             =   360
                  Width           =   2775
               End
               Begin VB.Label Label 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion a base de datos Sql Server"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   62
                  Left            =   1680
                  TabIndex        =   246
                  Top             =   150
                  Width           =   3105
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame Frame30 
               Caption         =   "Integracion Carvajal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   242
               Top             =   2295
               Width           =   4820
               Begin VB.CheckBox chkFEGesa 
                  Caption         =   "Grupo GESA"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   252
                  Top             =   240
                  Width           =   1740
               End
               Begin VB.CheckBox chkFECarbajal 
                  Caption         =   "InfoRest - Carvajal"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   243
                  Top             =   240
                  Width           =   1740
               End
            End
            Begin VB.Frame Frame25 
               Caption         =   "Integracion Paperless"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   960
               Left            =   120
               TabIndex        =   238
               Top             =   1320
               Width           =   4820
               Begin VB.TextBox TxtCorreoPape 
                  Enabled         =   0   'False
                  Height          =   350
                  Left            =   720
                  TabIndex        =   315
                  Top             =   500
                  Width           =   3975
               End
               Begin VB.CheckBox chkFEpape 
                  Caption         =   "Facturación Electrónica Paperless"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   240
                  Top             =   240
                  Width           =   2820
               End
               Begin VB.CheckBox chkFEubl21 
                  Caption         =   "Activa UBL 2.1"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   3240
                  TabIndex        =   239
                  Top             =   240
                  Width           =   1500
               End
               Begin VB.Label Label22 
                  Caption         =   "Correo:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   316
                  Top             =   550
                  Width           =   615
               End
               Begin VB.Label Label 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion TCP IP"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   68
                  Left            =   3075
                  TabIndex        =   241
                  Top             =   0
                  Width           =   1545
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame fraPaCarvajal 
               Caption         =   "Parametros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2415
               Left            =   120
               TabIndex        =   234
               Top             =   4800
               Width           =   4815
               Begin VB.TextBox txtParamCarv 
                  Height          =   1575
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   236
                  Top             =   240
                  Width           =   4575
               End
               Begin VB.TextBox txtCarvajalCorreos 
                  Height          =   405
                  Left            =   720
                  TabIndex        =   235
                  Top             =   1920
                  Width           =   3975
               End
               Begin VB.Label Label9 
                  Caption         =   "Correos:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   237
                  Top             =   2040
                  Width           =   735
               End
            End
            Begin VB.Frame Frame31 
               Caption         =   "Integracion TCI"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   232
               Top             =   2775
               Width           =   4820
               Begin VB.TextBox txtPOSTCI 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   3120
                  MaxLength       =   15
                  TabIndex        =   338
                  Top             =   150
                  Width           =   1455
               End
               Begin VB.CheckBox chkfeTCI 
                  Caption         =   "Facturación Electrónica TCI"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   233
                  Top             =   240
                  Width           =   3060
               End
            End
            Begin VB.Frame Frame32 
               Caption         =   "Integracion Bizlinks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   1080
               Left            =   120
               TabIndex        =   230
               Top             =   3240
               Width           =   4820
               Begin VB.CheckBox chklFEBizOffLine 
                  Caption         =   "Bizlinks OffLine"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   340
                  Top             =   300
                  Width           =   1455
               End
               Begin VB.CheckBox chkFEBiz 
                  Caption         =   "Facturación Electrónica Bizlinks"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   231
                  Top             =   300
                  Width           =   2580
               End
            End
            Begin VB.Frame Frame34 
               Caption         =   "Integracion Good Hope"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   228
               Top             =   4320
               Width           =   4820
               Begin VB.CheckBox chkFEGood 
                  Caption         =   "Facturación Electrónica Good Hope"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   229
                  Top             =   240
                  Width           =   3060
               End
            End
         End
         Begin VB.CheckBox chkInNC 
            Caption         =   "Incluir Notas de Credito en Liquidacion de Cajeros"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   214
            Top             =   2400
            Width           =   4215
         End
         Begin VB.CheckBox chkAnulacionNC 
            Caption         =   "Activar notas de credito por anulacion de documentos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   208
            Top             =   2160
            Width           =   4455
         End
         Begin VB.Frame Frame27 
            Caption         =   "Conexión a Base de Datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Left            =   120
            TabIndex        =   194
            Top             =   2760
            Width           =   4820
            Begin VB.CommandButton cmdValidar 
               Caption         =   "Validar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   2
               Left            =   3255
               Picture         =   "frmParametro.frx":03BA
               Style           =   1  'Graphical
               TabIndex        =   204
               Top             =   1230
               UseMaskColor    =   -1  'True
               Width           =   1410
            End
            Begin VB.TextBox txtClaveFE 
               Alignment       =   2  'Center
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
               IMEMode         =   3  'DISABLE
               Left            =   1410
               PasswordChar    =   "*"
               TabIndex        =   203
               Top             =   1560
               Width           =   1800
            End
            Begin VB.TextBox txtUsuarioFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   202
               Top             =   1245
               Width           =   1800
            End
            Begin VB.TextBox txtServidorFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   197
               Top             =   615
               Width           =   3250
            End
            Begin VB.TextBox txtBDFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   196
               Top             =   930
               Width           =   3250
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Clave :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   64
               Left            =   240
               TabIndex        =   201
               Top             =   1565
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Usuario :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   63
               Left            =   285
               TabIndex        =   200
               Top             =   1290
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Servidor :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   58
               Left            =   240
               TabIndex        =   199
               Top             =   660
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Base Datos :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   59
               Left            =   300
               TabIndex        =   198
               Top             =   975
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Motor BD: Sql Server"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   65
               Left            =   300
               TabIndex        =   195
               Top             =   285
               Width           =   1740
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame FrameTipoImpresion 
            Caption         =   "Tipo Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   3060
            TabIndex        =   189
            Top             =   210
            Width           =   1800
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código Barras"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   192
               Top             =   300
               Value           =   -1  'True
               Width           =   1485
            End
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código Hash"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   191
               Top             =   540
               Width           =   1275
            End
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código QR"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   190
               Top             =   800
               Width           =   1275
            End
         End
         Begin VB.CheckBox chkFacturacionE 
            Alignment       =   1  'Right Justify
            Caption         =   "Facturación Electrónica  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   188
            Top             =   345
            Width           =   2145
         End
         Begin VB.CheckBox chkAmbienteFE 
            Alignment       =   1  'Right Justify
            Caption         =   "Ambiente Producción  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   187
            Top             =   600
            Width           =   1980
         End
         Begin VB.TextBox txtCodigoFE 
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
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   186
            Top             =   945
            Width           =   750
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Maximo de dias para Anular un Documento:"
            Height          =   195
            Index           =   92
            Left            =   420
            TabIndex        =   351
            Top             =   1440
            Width           =   3090
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Código Facturación :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   50
            Left            =   570
            TabIndex        =   193
            Top             =   960
            Width           =   1515
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Cheff Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   182
         Top             =   3480
         Width           =   5535
         Begin VB.CheckBox chkCheffFiltroSubGrupo 
            Caption         =   "Permite el Filtro de Pedidos por Sub-Grupos."
            Height          =   285
            Left            =   240
            TabIndex        =   184
            Top             =   675
            Width           =   3975
         End
         Begin VB.CheckBox chkCheffFiltroSalon 
            Caption         =   "Permite el Filtro de Pedidos por Salones."
            Height          =   315
            Left            =   240
            TabIndex        =   183
            Top             =   300
            Width           =   3975
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Configuracion - Nota de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Index           =   17
         Left            =   120
         TabIndex        =   180
         Top             =   5520
         Width           =   3360
         Begin VB.CheckBox ChklBloquearNCsinEstado 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear Creacion de NC, sin estado Aceptado"
            Height          =   375
            Left            =   120
            TabIndex        =   330
            Top             =   1220
            Width           =   3090
         End
         Begin VB.CheckBox chkDesNCPG 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear NC como Forma de Pago :"
            Height          =   195
            Left            =   120
            TabIndex        =   213
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   960
            Width           =   3090
         End
         Begin VB.CheckBox chkNCElimina 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear Eliminar NC"
            Height          =   195
            Left            =   120
            TabIndex        =   210
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   720
            Width           =   3090
         End
         Begin VB.CheckBox chkNCParcial 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear NC Parciales"
            Height          =   195
            Left            =   120
            TabIndex        =   209
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   480
            Width           =   3090
         End
         Begin VB.CheckBox chkNCFecha 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquedo de Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   181
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   3090
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Configuracion -  BAR"
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
         Index           =   8
         Left            =   6840
         TabIndex        =   172
         Top             =   2140
         Width           =   3525
         Begin VB.CommandButton cmdbuscarItemCover 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2840
            Picture         =   "frmParametro.frx":07FC
            Style           =   1  'Graphical
            TabIndex        =   223
            Top             =   1080
            Width           =   630
         End
         Begin VB.TextBox txtCodigoItemCover 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   177
            Tag             =   "02155454555"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtMontoMinCover 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   175
            Top             =   700
            Width           =   1095
         End
         Begin VB.CheckBox chkCover 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar cargo automático de Cover a Pedido"
            Height          =   375
            Left            =   120
            TabIndex        =   173
            Top             =   240
            Width           =   3285
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Codigo de Item Cover"
            Height          =   195
            Index           =   89
            Left            =   120
            TabIndex        =   176
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Label Label 
            Caption         =   "Monto Minimo / PAX (Cliente)"
            Height          =   255
            Index           =   90
            Left            =   120
            TabIndex        =   174
            Top             =   720
            Width           =   2175
         End
      End
      Begin VB.Frame FrmVisor 
         Caption         =   "Visor de 8"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   167
         Top             =   4680
         Width           =   5535
         Begin VB.TextBox txtvisortiempo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   170
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chktactil 
            Alignment       =   1  'Right Justify
            Caption         =   "Visor Tactil"
            Height          =   195
            Left            =   4080
            TabIndex        =   169
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkVisor8 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar Visor de 8"" (AMC)"
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   255
            Width           =   2295
         End
         Begin VB.Label Label 
            Caption         =   "Tiempo de espera de mensaje e Inactividad. (Segundos) :"
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   171
            Top             =   600
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Transferencia Gratuita"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   7
         Left            =   3600
         TabIndex        =   164
         Top             =   3180
         Width           =   3165
         Begin VB.TextBox txtGlosaImpresion 
            Height          =   495
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   179
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox TxtCuentaContable 
            Height          =   285
            Left            =   1440
            TabIndex        =   166
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Glosa Impresion:"
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta Contable:"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Configuración Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   4
         Left            =   3600
         TabIndex        =   152
         Top             =   1500
         Width           =   3165
         Begin VB.TextBox txtCodigoMarca 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   160
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdBusca 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            Picture         =   "frmParametro.frx":08FE
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   1200
            Width           =   630
         End
         Begin VB.TextBox txtCodigoUbigeo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   155
            Top             =   1200
            Width           =   990
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   154
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtCodigoTienda 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   153
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Marca :"
            Height          =   195
            Index           =   56
            Left            =   0
            TabIndex        =   161
            Top             =   640
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ubigeo :"
            Height          =   210
            Index           =   55
            Left            =   720
            TabIndex        =   158
            Top             =   1200
            Width           =   645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Tienda :"
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   157
            Top             =   920
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código Empresa :"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   156
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Activación Club"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3600
         TabIndex        =   145
         Top             =   5880
         Width           =   3165
         Begin VB.CheckBox chkClub 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa Club :"
            Height          =   195
            Left            =   120
            TabIndex        =   146
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Anfitrionas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2925
         Left            =   -74880
         TabIndex        =   128
         Top             =   435
         Width           =   5565
         Begin VB.CheckBox chkAgradecimiento 
            Caption         =   "Enviar Email de Agradecimiento por Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   142
            Top             =   5220
            Visible         =   0   'False
            Width           =   4440
         End
         Begin VB.CheckBox chkRecordatorio 
            Caption         =   "Enviar Email de Recordatorio de Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   141
            Top             =   2970
            Visible         =   0   'False
            Width           =   4440
         End
         Begin VB.CheckBox chkConfirmacion 
            Caption         =   "Enviar Email de Confirmación de Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   140
            Top             =   600
            Width           =   4440
         End
         Begin VB.TextBox txtAgradecimiento 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   139
            Top             =   5445
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.TextBox txtRecordatorio 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   138
            Top             =   3195
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.TextBox txtConfirmacion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   825
            Width           =   5385
         End
         Begin VB.TextBox txtToleranciaReserva 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4275
            MaxLength       =   3
            TabIndex        =   135
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Min Toleracia Reserva :"
            Height          =   210
            Index           =   52
            Left            =   1650
            TabIndex        =   136
            Top             =   255
            Width           =   2235
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Central de Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2130
         Left            =   -69240
         TabIndex        =   120
         Top             =   1905
         Width           =   4545
         Begin VB.TextBox txtMaxMotorizado 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2925
            MaxLength       =   5
            TabIndex        =   219
            Top             =   1480
            Width           =   1320
         End
         Begin VB.TextBox txtAsignacionMotorizado 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   124
            Top             =   885
            Width           =   1320
         End
         Begin VB.TextBox txtTiempoDelivery 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2925
            TabIndex        =   123
            Top             =   555
            Width           =   1320
         End
         Begin VB.TextBox txtDiaDelivery 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2925
            TabIndex        =   122
            Top             =   225
            Width           =   1320
         End
         Begin VB.CheckBox chkHoraEntrega 
            Alignment       =   1  'Right Justify
            Caption         =   "Asignar hora de entrega desde Despachador"
            Height          =   255
            Left            =   225
            TabIndex        =   121
            Top             =   1200
            Width           =   4035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Maximo de Motorizados a Asignar:"
            Height          =   195
            Index           =   69
            Left            =   240
            TabIndex        =   220
            Top             =   1560
            Width           =   2625
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Asignación a Motorizados S/."
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   127
            Top             =   930
            Width           =   2100
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Días de búsqueda Delivery :"
            Height          =   195
            Index           =   42
            Left            =   450
            TabIndex        =   126
            Top             =   270
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Min promedio de entrega del :"
            Height          =   195
            Index           =   45
            Left            =   510
            TabIndex        =   125
            Top             =   600
            Width           =   2190
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmMobile 
         Caption         =   "Mobile Inforest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1380
         Left            =   -69225
         TabIndex        =   117
         Top             =   435
         Width           =   4545
         Begin VB.CheckBox chkMCCaja 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa solicitud de Autorización al  cambiar de Caja"
            Height          =   495
            Left            =   240
            TabIndex        =   119
            Top             =   840
            Width           =   3975
         End
         Begin VB.CheckBox chkMUnidadNegocio 
            Alignment       =   1  'Right Justify
            Caption         =   "Permite el Filtro por Unidad de Negocio en los dispositivos Móviles"
            Height          =   495
            Left            =   240
            TabIndex        =   118
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Equipos Biométricos (Huella Dactilar)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   6840
         TabIndex        =   111
         Top             =   1350
         Width           =   3525
         Begin VB.CheckBox chkSecugen 
            Alignment       =   1  'Right Justify
            Caption         =   "Hamster Plus (SecuGen) :      "
            Height          =   255
            Left            =   240
            TabIndex        =   113
            ToolTipText     =   "Modelo HSDU03P"
            Top             =   480
            Width           =   2565
         End
         Begin VB.CheckBox chkDigital 
            Alignment       =   1  'Right Justify
            Caption         =   "Digital Persona 4500  :"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Configuración de Factura Variable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   18
         Left            =   120
         TabIndex        =   105
         Top             =   4200
         Width           =   3360
         Begin VB.TextBox txtCabeceraV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   134
            Top             =   520
            Width           =   1470
         End
         Begin VB.TextBox txtItemV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   133
            Top             =   240
            Width           =   1470
         End
         Begin VB.TextBox txtPieV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   132
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Pie de Pag :"
            Height          =   210
            Index           =   41
            Left            =   150
            TabIndex        =   108
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   46
            Left            =   180
            TabIndex        =   107
            Top             =   560
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   47
            Left            =   255
            TabIndex        =   106
            Top             =   280
            Width           =   1410
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Varios "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Index           =   1
         Left            =   3600
         TabIndex        =   32
         Top             =   420
         Width           =   3165
         Begin VB.TextBox txtDia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   101
            Top             =   570
            Width           =   1470
         End
         Begin VB.TextBox txtCorrelativo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   100
            Top             =   180
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Días en la grilla :"
            Height          =   195
            Index           =   40
            Left            =   45
            TabIndex        =   103
            Top             =   585
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Correlativo :"
            Height          =   195
            Index           =   29
            Left            =   450
            TabIndex        =   102
            Top             =   225
            Width           =   840
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Configuración de Factura Manual "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   3360
         Begin VB.TextBox txtItem 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   131
            Top             =   225
            Width           =   1470
         End
         Begin VB.TextBox txtCabecera 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   130
            Top             =   525
            Width           =   1470
         End
         Begin VB.TextBox txtDetalle 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   129
            Top             =   825
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Detalle :"
            Height          =   210
            Index           =   17
            Left            =   255
            TabIndex        =   99
            Top             =   885
            Width           =   1410
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   15
            Left            =   180
            TabIndex        =   98
            Top             =   585
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   1
            Left            =   255
            TabIndex        =   97
            Top             =   285
            Width           =   1410
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Puntos  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   2
         Left            =   6840
         TabIndex        =   33
         Top             =   420
         Width           =   3525
         Begin VB.TextBox txtClub 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   96
            Top             =   210
            Width           =   1470
         End
         Begin VB.TextBox txtPunto 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   95
            Top             =   520
            Width           =   1470
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Puntos :"
            Height          =   195
            Index           =   38
            Left            =   300
            TabIndex        =   151
            Top             =   255
            Width           =   1470
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor de un Punto S/. :"
            Height          =   195
            Index           =   37
            Left            =   135
            TabIndex        =   150
            Top             =   560
            Width           =   1635
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Dia Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   3360
         Begin VB.OptionButton optDCAutomatico 
            Caption         =   "Automático / Hora de Cierre"
            Height          =   465
            Left            =   195
            TabIndex        =   93
            Top             =   240
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optDCManual 
            Caption         =   "Manual (Cierre de Turno)"
            Height          =   195
            Left            =   180
            TabIndex        =   92
            Top             =   740
            Width           =   2055
         End
         Begin VB.CheckBox chkImprimeDiaContable 
            Caption         =   "Impresión de Dia Contable en Documentos"
            Height          =   435
            Left            =   195
            TabIndex        =   91
            Top             =   940
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker dtpHoraDC 
            Height          =   315
            Left            =   1800
            TabIndex        =   94
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   16842755
            UpDown          =   -1  'True
            CurrentDate     =   38587.2083333333
         End
      End
      Begin VB.Frame Frame 
         Caption         =   " Configuración de Guía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   3360
         Begin VB.TextBox txtItemGuia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1815
            TabIndex        =   87
            Top             =   210
            Width           =   1470
         End
         Begin VB.TextBox txtCabeceraGuia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   86
            Top             =   510
            Width           =   1470
         End
         Begin VB.TextBox txtDetalleGuia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   85
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   34
            Left            =   180
            TabIndex        =   89
            Top             =   540
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   35
            Left            =   270
            TabIndex        =   90
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Detalle :"
            Height          =   210
            Index           =   33
            Left            =   255
            TabIndex        =   88
            Top             =   840
            Width           =   1410
         End
      End
      Begin VB.Frame fraRuc 
         Caption         =   "Identificador Tributario "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -67920
         TabIndex        =   29
         Top             =   7440
         Width           =   3345
         Begin VB.TextBox txtLongitud 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2550
            TabIndex        =   82
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton opcLongitud 
            Caption         =   " =   Longitud"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   81
            Top             =   540
            Width           =   1470
         End
         Begin VB.OptionButton opcLongitud 
            Caption         =   " > = Longitud"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   80
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Longitud Identificador Tributario"
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   2235
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Longitud"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   83
            Top             =   525
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Impuestos y Extras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   -74880
         TabIndex        =   27
         Top             =   6720
         Width           =   6855
         Begin VB.CheckBox chkVisibleimpuesto 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocultar porcentaje en impresiòn de documento (impuesto 1)"
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   359
            Top             =   360
            Visible         =   0   'False
            Width           =   2600
         End
         Begin VB.CheckBox chkVisibleimpuesto 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocultar porcentaje en impresiòn de documento (impuesto 2)"
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   358
            Top             =   820
            Width           =   2600
         End
         Begin VB.CheckBox chkVisibleimpuesto 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocultar porcentaje en impresiòn de documento (impuesto 3)"
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   357
            Top             =   1280
            Width           =   2600
         End
         Begin VB.TextBox txtIImp3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3210
            TabIndex        =   24
            Top             =   1275
            Width           =   690
         End
         Begin VB.TextBox txtIImp2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3210
            TabIndex        =   22
            Top             =   840
            Width           =   690
         End
         Begin VB.TextBox txtIImp1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3210
            TabIndex        =   20
            Top             =   405
            Width           =   690
         End
         Begin VB.TextBox txtDImp3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   975
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1275
            Width           =   2200
         End
         Begin VB.TextBox txtDImp2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   975
            MaxLength       =   50
            TabIndex        =   21
            Top             =   840
            Width           =   2200
         End
         Begin VB.TextBox txtDImp1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   975
            MaxLength       =   50
            TabIndex        =   19
            Top             =   480
            Width           =   2200
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   3195
            TabIndex        =   79
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   1680
            TabIndex        =   78
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Extra 2 :"
            Height          =   195
            Index           =   28
            Left            =   330
            TabIndex        =   77
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Extra 1 :"
            Height          =   195
            Index           =   27
            Left            =   330
            TabIndex        =   76
            Top             =   885
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto :"
            Height          =   195
            Index           =   26
            Left            =   195
            TabIndex        =   75
            Top             =   450
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   23
            Left            =   3960
            TabIndex        =   74
            Top             =   1290
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   22
            Left            =   3960
            TabIndex        =   73
            Top             =   855
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   21
            Left            =   3960
            TabIndex        =   72
            Top             =   420
            Width           =   210
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Activaciones "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7035
         Left            =   -66840
         TabIndex        =   26
         Top             =   360
         Width           =   2250
         Begin VB.CheckBox chkAlertaVenta 
            Caption         =   "Alerta Descargo de Venta"
            Height          =   375
            Left            =   105
            TabIndex        =   354
            Top             =   6600
            Width           =   1935
         End
         Begin VB.CheckBox chklActivarIncrementoPrecios 
            Caption         =   "Programar incremento de precios de los productos"
            Height          =   495
            Left            =   80
            TabIndex        =   341
            Top             =   5520
            Width           =   2130
         End
         Begin VB.CheckBox chkValidaDNI 
            Caption         =   "Validar DNI"
            Height          =   255
            Left            =   70
            TabIndex        =   212
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CheckBox chkTCenImp 
            Caption         =   "Ver T.C. en Imp."
            Height          =   255
            Left            =   70
            TabIndex        =   211
            Top             =   5280
            Width           =   1815
         End
         Begin VB.CheckBox ChkActCuentaCorriente 
            Caption         =   "Act. de Cuentas C."
            Height          =   315
            Left            =   75
            TabIndex        =   163
            Top             =   6000
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox chkPagoAntesImpresion 
            Caption         =   "Registro de Pago antes de emisión de Comprobantes"
            Height          =   675
            Left            =   75
            TabIndex        =   162
            Top             =   4420
            Width           =   1935
         End
         Begin VB.CheckBox chkEventos 
            Caption         =   "Enlace Eventos"
            Height          =   195
            Left            =   75
            TabIndex        =   149
            Top             =   1150
            Width           =   1650
         End
         Begin VB.CheckBox chkTarjeta 
            Caption         =   "Validación Tarjeta"
            Height          =   195
            Left            =   75
            TabIndex        =   147
            Top             =   2100
            Width           =   1665
         End
         Begin VB.CheckBox chkControlEnviosProduccion 
            Caption         =   "Activa envios a producción por usuario"
            Height          =   435
            Left            =   75
            TabIndex        =   144
            Top             =   3960
            Width           =   1935
         End
         Begin VB.CheckBox chkEnvioAutomatico 
            Caption         =   "Envio a producción automático"
            Height          =   435
            Left            =   75
            TabIndex        =   143
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox chkControlUsuario 
            Caption         =   "Control de Usuarios Por Nivel"
            Height          =   435
            Left            =   75
            TabIndex        =   110
            Top             =   3000
            Width           =   1815
         End
         Begin VB.CheckBox chkConsultaDescargo 
            Caption         =   "Activa Consulta de Descargo de Venta al Cierre de Turno"
            Height          =   615
            Left            =   75
            TabIndex        =   104
            Top             =   2350
            Width           =   2025
         End
         Begin VB.CheckBox chkMultiLocal 
            Caption         =   "Enlace Multilocal"
            Height          =   195
            Left            =   75
            TabIndex        =   71
            Top             =   1860
            Width           =   1665
         End
         Begin VB.CheckBox chkComboGeneral 
            Caption         =   "Listado General(Combos)"
            Height          =   300
            Left            =   75
            TabIndex        =   70
            Top             =   1580
            Width           =   2070
         End
         Begin VB.CheckBox ChkEquivalencia 
            Caption         =   "Muestra Equivalencia"
            Height          =   195
            Left            =   75
            TabIndex        =   69
            Top             =   1380
            Width           =   1890
         End
         Begin VB.CheckBox chkCierre 
            Caption         =   "Cierre a Ciegas"
            Height          =   195
            Left            =   75
            TabIndex        =   68
            Top             =   210
            Width           =   1365
         End
         Begin VB.CheckBox chkInfhotel 
            Caption         =   "Enlace Infhotel"
            Height          =   195
            Left            =   75
            TabIndex        =   67
            Top             =   930
            Width           =   1410
         End
         Begin VB.CheckBox chkAlmacen 
            Caption         =   "Enlace Almacén"
            Height          =   195
            Left            =   75
            TabIndex        =   66
            Top             =   690
            Width           =   1590
         End
         Begin VB.CheckBox chkPrinter 
            Caption         =   "Kitchen Printer"
            Height          =   195
            Left            =   75
            TabIndex        =   65
            Top             =   450
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6225
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   7965
         Begin VB.TextBox txtCabPreCuenta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            TabIndex        =   355
            Top             =   5740
            Width           =   5385
         End
         Begin VB.TextBox txtMonedaE3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6495
            MaxLength       =   30
            TabIndex        =   259
            Top             =   3075
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtMonE3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5520
            MaxLength       =   3
            TabIndex        =   258
            Top             =   3075
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtMonE2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5520
            MaxLength       =   3
            TabIndex        =   257
            Top             =   2760
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtMonedaE2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6495
            MaxLength       =   30
            TabIndex        =   256
            Top             =   2760
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtFax 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5680
            MaxLength       =   30
            TabIndex        =   116
            Top             =   1470
            Width           =   2175
         End
         Begin VB.TextBox txtRetencion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2475
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   5040
            Width           =   5385
         End
         Begin VB.TextBox txtPieFE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   4680
            Width           =   5385
         End
         Begin VB.TextBox txtPie 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   4050
            Width           =   5385
         End
         Begin VB.TextBox txtDireccion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   150
            TabIndex        =   2
            Top             =   820
            Width           =   5385
         End
         Begin VB.TextBox txtEmail 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1785
            Width           =   3345
         End
         Begin VB.TextBox txtMonedaE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3450
            MaxLength       =   30
            TabIndex        =   11
            Top             =   3075
            Width           =   1395
         End
         Begin VB.TextBox txtMonE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   3
            TabIndex        =   10
            Top             =   3075
            Width           =   915
         End
         Begin VB.TextBox txtTelefono 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1480
            Width           =   2385
         End
         Begin VB.TextBox txtRUC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   15
            TabIndex        =   7
            Top             =   2430
            Width           =   3345
         End
         Begin VB.TextBox txtWebPage 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2100
            Width           =   5385
         End
         Begin VB.TextBox txtSocial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   1
            Top             =   510
            Width           =   5385
         End
         Begin VB.TextBox txtMonN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   3
            TabIndex        =   8
            Top             =   2760
            Width           =   915
         End
         Begin VB.TextBox txtMonedaN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3450
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2760
            Width           =   1395
         End
         Begin VB.TextBox txtComercial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   0
            Top             =   200
            Width           =   5385
         End
         Begin VB.TextBox txtPiePreCuenta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   3735
            Width           =   5385
         End
         Begin VB.TextBox txtElimina 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   15
            TabIndex        =   16
            Top             =   4380
            Width           =   5385
         End
         Begin VB.TextBox txtContribuyenteEspecial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   12
            Top             =   3405
            Width           =   915
         End
         Begin VB.TextBox txtDireccion2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2475
            MaxLength       =   150
            TabIndex        =   3
            Top             =   1160
            Width           =   5385
         End
         Begin MSComCtl2.DTPicker dtpContribuyenteEspecial 
            Height          =   315
            Left            =   3450
            TabIndex        =   13
            Top             =   3390
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
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
            Format          =   16842753
            UpDown          =   -1  'True
            CurrentDate     =   2.20833333333333
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Cab. de la Precuenta:"
            Height          =   375
            Left            =   75
            TabIndex        =   356
            Top             =   5760
            Width           =   2295
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "ME 3:"
            Height          =   210
            Index           =   67
            Left            =   4920
            TabIndex        =   255
            Top             =   3060
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "ME 2:"
            Height          =   210
            Index           =   66
            Left            =   4920
            TabIndex        =   254
            Top             =   2760
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax  :"
            Height          =   210
            Index           =   51
            Left            =   4920
            TabIndex        =   115
            Top             =   1500
            Width           =   645
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto para Agentes de Retención :"
            Height          =   450
            Index           =   49
            Left            =   175
            TabIndex        =   114
            Top             =   5160
            Width           =   2205
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie Facturación :"
            Height          =   210
            Index           =   48
            Left            =   195
            TabIndex        =   109
            Top             =   4695
            Width           =   2205
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie del Documento :"
            Height          =   210
            Index           =   10
            Left            =   75
            TabIndex        =   64
            Top             =   4095
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda Extranjera :"
            Height          =   210
            Index           =   9
            Left            =   75
            TabIndex        =   63
            Top             =   3120
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "WebPage :"
            Height          =   210
            Index           =   6
            Left            =   75
            TabIndex        =   62
            Top             =   2145
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "E-Mail :"
            Height          =   210
            Index           =   5
            Left            =   75
            TabIndex        =   61
            Top             =   1815
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Teléfonos  :"
            Height          =   210
            Index           =   4
            Left            =   75
            TabIndex        =   60
            Top             =   1485
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Razón Social (Bussines Name) :"
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   59
            Top             =   560
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda Nacional :"
            Height          =   210
            Index           =   8
            Left            =   75
            TabIndex        =   58
            Top             =   2790
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Id. Tributario (Federal Id.) :"
            Height          =   210
            Index           =   7
            Left            =   75
            TabIndex        =   57
            Top             =   2460
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Dirección :"
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   56
            Top             =   870
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Razón Comercial (Legal Name) :"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   55
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie de la Precuenta :"
            Height          =   210
            Index           =   36
            Left            =   75
            TabIndex        =   54
            Top             =   3765
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mensaje de Eliminación Pedido :"
            Height          =   210
            Index           =   11
            Left            =   75
            TabIndex        =   53
            Top             =   4410
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Ley Contribuyente Especial :"
            Height          =   210
            Index           =   39
            Left            =   75
            TabIndex        =   52
            Top             =   3435
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Height          =   210
            Index           =   30
            Left            =   75
            TabIndex        =   51
            Top             =   1180
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Porcentajes por T/P "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   3360
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   2280
      Begin VB.TextBox txtCanal5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   38
         Top             =   1185
         Width           =   720
      End
      Begin VB.TextBox txtCanal4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   37
         Top             =   870
         Width           =   720
      End
      Begin VB.TextBox txtllevar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   36
         Top             =   540
         Width           =   720
      End
      Begin VB.TextBox txtDelivery 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   35
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 5:"
         Height          =   195
         Index           =   44
         Left            =   135
         TabIndex        =   49
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 4:"
         Height          =   195
         Index           =   43
         Left            =   135
         TabIndex        =   48
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 3:"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   47
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   2025
         TabIndex        =   46
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   2025
         TabIndex        =   45
         Top             =   885
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   2025
         TabIndex        =   44
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   2025
         TabIndex        =   43
         Top             =   570
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 2:"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   42
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7920
      Picture         =   "frmParametro.frx":0A00
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8760
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   9240
      Picture         =   "frmParametro.frx":0B02
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8760
      Width           =   1170
   End
   Begin VB.Label Label20 
      Caption         =   "Glosa Impresion Inafecto:"
      Height          =   495
      Left            =   0
      TabIndex        =   312
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsParametro As Recordset
Dim RsProducto As Recordset
Dim lImprimeCodigoBarras As Boolean

'KDS
Private Sub btnKDS_Click()
    'frmKDSConfiguracion.Show vbModal
    frmConfiguracionParametro.Show vbModal
End Sub

Private Sub chkActivaAnticipo_Click()
    If bCargaFormulario Then Exit Sub
    
    If Me.chkActivaAnticipo.Value = 1 Then
        MsgBox "Tener Presente:" + vbNewLine + "Activar Anticipos puede causar problemas con las interfaces contables, Favor de Contactarse con Infhotel Servicios Informaticos S.A.C!!!"
    End If
End Sub

Private Sub chkAgradecimiento_Click()
 If chkAgradecimiento.Value = 1 Then
        Me.txtAgradecimiento.Enabled = True
 Else
    txtAgradecimiento.Enabled = True
 End If
End Sub

Private Sub chkConfirmacion_Click()
 If chkConfirmacion.Value = 1 Then
        txtConfirmacion.Enabled = True
 Else
    txtConfirmacion.Enabled = True
 End If
End Sub

Private Sub chkDigital_Click()
If chkDigital.Value = 1 Then
    chkSecugen.Value = 0
End If
End Sub

Private Sub chkFEGesa_Click()
    If Me.chkFEGesa.Value = 1 Then
        Me.chkfeTCI.Value = 0
        chkFEOfisis.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFESpring.Value = 0
        chkFECarbajal.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        chkFacturacionE.Value = 1
        Me.fraPaCarvajal.Visible = False
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGood.Value = 0
    End If

End Sub

Private Sub chkFEGood_Click()
   If Me.chkFEGood.Value = 1 Then
        Me.chkfeTCI.Value = 0
        chkFEOfisis.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFESpring.Value = 0
        chkFECarbajal.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        chkFacturacionE.Value = 1
        Me.fraPaCarvajal.Visible = True
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    Else
        Me.fraPaCarvajal.Visible = False
    End If
End Sub
Private Sub chkFEBiz_Click()
   If Me.chkFEBiz.Value = 1 Then
        Me.chkfeTCI.Value = 0
        Me.chkFEGood.Value = 0
        chkFEOfisis.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFESpring.Value = 0
        chkFECarbajal.Value = 0
        chkFacturacionE.Value = 1
        
        Me.fraPaCarvajal.Visible = True
        Me.txtNumdiasAnulacion.Enabled = True
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    Else
        Me.fraPaCarvajal.Visible = False
        Me.txtNumdiasAnulacion.Enabled = False
    End If
End Sub

Private Sub chkFECarbajal_Click()
    If chkFECarbajal.Value = 1 Then
        chkFEOfisis.Value = 0
        Me.chkFEGood.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFESpring.Value = 0
        chkfeTCI.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        chkFacturacionE.Value = 1
        Me.fraPaCarvajal.Visible = True
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    Else
        Me.fraPaCarvajal.Visible = False
    End If
End Sub
Private Sub chkFETCI_Click()
    If chkfeTCI.Value = 1 Then
        chkFEOfisis.Value = 0
        Me.chkFEGood.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFESpring.Value = 0
        chkFECarbajal.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        Me.fraPaCarvajal.Visible = False
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    End If
End Sub
Private Sub chkFEOfisis_Click()
    If chkFEOfisis.Value = 1 Then
        chkFESpring.Value = 0
        Me.chkFEGood.Value = 0
        chkFEpape.Value = 0
        chkFECarbajal.Value = 0
        chkfeTCI.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        chkFacturacionE.Value = 1
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    End If
End Sub

Private Sub chkFEpape_Click()
    If chkFEpape.Value = 1 Then
        chkFESpring.Value = 0
        Me.chkFEGood.Value = 0
        chkFEOfisis.Value = 0
        chkfeTCI.Value = 0
        chkFEGesa.Value = 0
        'chkInNC.value = 0
        chkFECarbajal.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        
        If optOpcion(0).Value = True Then
            optOpcion(0).Value = False
            If optOpcion(1).Value = False Then
                If optOpcion(2).Value = False Then
                    optOpcion(1).Value = True
                End If
            Else
            
            End If
        End If
        optOpcion(0).Visible = False
        chkFacturacionE.Value = 1
        Me.chkFEubl21.Enabled = True
        Me.TxtCorreoPape.Enabled = True
        
    Else
        optOpcion(0).Visible = True
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
    End If
End Sub

Private Sub chkFESpring_Click()
    If chkFESpring.Value = 1 Then
        chkFEOfisis.Value = 0
        Me.chkFEGood.Value = 0
        'chkInNC.value = 0
        chkFEpape.Value = 0
        chkFECarbajal.Value = 0
        chkfeTCI.Value = 0
        Me.chkFEBiz.Value = 0
        Me.chklFEBizOffLine.Value = 0
        Me.txtNumdiasAnulacion.Enabled = False
        chkFacturacionE.Value = 1
        Me.chkFEubl21.Enabled = False
        Me.TxtCorreoPape.Enabled = False
        Me.chkFEubl21.Value = 0
        chkFEGesa.Value = 0
    End If
End Sub


Private Sub chkRecordatorio_Click()
 If chkRecordatorio.Value = 1 Then
    Me.txtRecordatorio.Enabled = True
 Else
    txtRecordatorio.Enabled = True
 End If
End Sub

Private Sub ChkSAP_Click()

    If ChkSAP.Value = 1 Then
        Frame20.Visible = True
    Else
        Frame20.Visible = False
    End If

End Sub

Private Sub chkSecugen_Click()
If chkSecugen.Value = 1 Then
    chkDigital.Value = 0
End If
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigo as Codigo, tDistrito as Descripcion, tProvincia as Provincia, tDepartamento as Departamento from TUBIGEO order by tCodigo asc"
   
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 1
   Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Distrito", 2, "Descripcion", 1500, 0, 0, "", _
                                          "Provincia", 2, "Provincia", 2500, 0, 0, "", _
                                          "Departamento", 2, "Departamento", 3000, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   txtCodigoUbigeo.Text = sCodigo
End Sub

Private Sub cmdBuscaAnticipo_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
        frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoAnticipo.Text = sCodigo
       End If

   ' End If

Exit Sub
fin:
MsgBox (Error)
End Sub

Private Sub cmdbuscarItemCover_Click()
On Error GoTo fin
 
   If IIf(chkCover.Value, 1, 0) = 1 Then
        Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
        frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoItemCover.Text = sCodigo
       Else
            txtCodigoItemCover.Text = ""
       End If

    End If

Exit Sub
fin:
MsgBox (Error)
End Sub

Private Sub cmdBuscaTarjeta_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
       frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoTarjeta.Text = sCodigo
       End If

   ' End If

Exit Sub
fin:
MsgBox (Error)
End Sub

Private Sub cmdItemDelivery_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
       frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            Me.txtItemDelivery.Text = sCodigo
       End If

   ' End If

Exit Sub
fin:
MsgBox (Error)
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  
If TxtlRetencion.Text = "" Then
    TxtlRetencion.Text = 0
End If


If TxtMontoMinRete.Text = "" Or TxtMontoMinRete.Text = "0" Then
    TxtMontoMinRete.Text = 700
End If
  
    If Index = 0 Then
    '--- Sap ------
     If ChkSAP Then
         If TxtCodAlmcSAP.Text = "" Then MsgBox "Ingrese el codigo de tienda.", vbExclamation, sMensaje: TxtCodAlmcSAP.SetFocus: Exit Sub
    End If
    
    Dim xImprimeCodigoBarras As Integer
    If optOpcion(0).Value = True Then
         xImprimeCodigoBarras = 1
    Else
         xImprimeCodigoBarras = 0
    End If
    
    If chkFEOfisis.Value = 1 And optOpcion(0).Value = True Then
        MsgBox "Tipo Impresión (Ofisis): Habilitada unicamente para Hash o QR.", vbExclamation, "Integración Ofisis": optOpcion(1).SetFocus: Exit Sub
    End If
    
    If chkFESpring.Value = 1 Then
        If optOpcion(0).Value = True Then
            MsgBox "Tipo Impresión (Spring): Habilitada unicamente para Hash y QR.", vbExclamation, "Integración Spring": optOpcion(1).SetFocus: Exit Sub
        End If
    End If
    
    If chkFECarbajal.Value = 1 Then
        If optOpcion(0).Value = True Or optOpcion(1).Value = True Then
            MsgBox "Tipo Impresión (Carvajal): Habilitada unicamente para QR.", vbExclamation, "Integración Carvajal": optOpcion(2).SetFocus: Exit Sub
        End If
    End If
    
    If chkFEOfisis.Value = 1 Or chkFESpring.Value = 1 Or chkFEpape.Value = 1 Or chkFECarbajal.Value = 1 Then
        chkFacturacionE.Value = 1
    End If
   
   If TxtlRetencion.Text > 100 Then
        MsgBox "Ingresar un dato valido en el porcentaje de retención", vbExclamation, "Porcentaje de Retención"
        Exit Sub
   End If
   
   
      Screen.MousePointer = vbHourglass
      'KDS
      
           sPasa = txtRUC.Text
            'Inserta Movimiento auditoria
            lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TPARAMETRO", "PARAMETRO", "02", sUsuario, sPasa, "", _
                "tidentificacionTributaria", "Id Tributario", Me.txtRUC.Text, "tRazonSocial", "Razon Social", txtSocial.Text, "tRazonComercial", "Razon Comercial", txtComercial.Text, "tDireccion", "Direccion", txtDireccion.Text, "tTelefono", "Telefono", txtTelefono.Text, "tEmail", "Email", txtEmail.Text, "tWebPage", "Pagina Web", txtWebPage.Text, "tDireccion2", "Direccion2", txtDireccion2.Text, _
                "tMonN", "Simb Moneda Nacional", txtMonN.Text, "tMonedaN", "Desc Moneda Nacional", txtMonedaN.Text, "tMonE", "Simb Moneda Extranjera", txtMonE.Text, "tMonedaE", "Desc Moneda Extranjera", txtMonedaE.Text, "tMonE2", "Simb Moneda Extranjera2", txtMonE2.Text, "tMonedaE2", "Desc Moneda Extranjera2", txtMonedaE2.Text, "tMonE3", "Simb Moneda Extranjera3", txtMonE3.Text, _
                "tMonedaE3", "Desc Moneda Extranjera3", txtMonedaE3.Text, "tPiePrecuenta", "Pie Precuenta", txtPiePreCuenta.Text, "tPie", "Pie Documentos", txtPie.Text, "tElimina", "Glosa de Eliminacion", txtElimina.Text, "tImpuesto1", "Descripcion Impuesto 1", txtDImp1.Text, "Impuesto1", "Porcentaje Impuesto 1", val(txtIImp1.Text), "tImpuesto2", "Descripcion Impuesto 2", txtDImp2.Text, _
                "Impuesto2", "Porcentaje Impuesto 2", val(txtIImp2.Text), "tImpuesto3", "Descripcion Impuesto 3", txtDImp3.Text, "Impuesto3", "Porcentaje Impuesto 3", val(txtIImp3.Text), "nDelivery", "Porc Recargo por Tipo de Pedido 2", val(txtDelivery.Text), "nLlevar", "Porc Recargo por Tipo de Pedido 3", val(txtllevar.Text), "nCanal4", "Porc Recargo por Tipo de Pedido 4", val(txtCanal4.Text), _
                "nCanal5", "Porc Recargo por Tipo de Pedido 5", val(txtCanal5.Text), "nCorrelativo", "Correlativo de Pedidos", val(txtCorrelativo.Text), "nItem", "Impresion Factura Manual Items", txtItem.Text, "nCabecera", "Impresion Factura Manual Cabecera", txtCabecera.Text, "nDetalle", "Impresion Factura Manual Detalle", txtDetalle.Text, "nItemGuia", "Impresion Guia Items", txtItemGuia.Text, "nCabeceraGuia", "Impresion Guia Cabecera", txtCabeceraGuia.Text, "nDetalleGuia", "Impresion Guia Detalle", txtDetalleGuia.Text, _
                "nDias", "Dias en Grilla", txtDia.Text, "nDiasDelivery", "Dias Busqueda Delivery", txtDiaDelivery.Text, "nTiempoMinutoCD", "Tiempo Entrega Delivery", txtTiempoDelivery.Text, "nAsignacionMotorizado", "Monto Maximo Motorizado", val(txtAsignacionMotorizado.Text), _
                "nLongitud", "Longitud Identificador Tributario", val(txtLongitud.Text), "lLongitud", "Condicion de Longitud", IIf(opcLongitud(0).Value, "Verdadero", "Falso"), "tClub", "Nombre de Punto Club", txtClub.Text, "nPunto", "Valor de Punto Club", val(txtPunto.Text), _
                "lDiaContableAutomatico", "Flag Dia Contable Automatico", IIf(optDCAutomatico.Value, "Verdadero", "Falso"), "lDiaContableManual", "Flag Dia Contable Manual", IIf(optDCManual.Value, "Verdadero", "Falso"), "thoracierrediacontable", "Dia Contable Hora Cierre Automatico", Format(Me.dtpHoraDC.Value, "HH:mm"), "lImprimeDiacontable", "Dia Contable Impresion Documentos", IIf(Me.chkImprimeDiaContable.Value, "Verdadero", "Falso"), _
                "lKds", "Flag Kds", IIf(frmConfiguracionParametro.chkKDS.Value, "Verdadero", "Falso"), "lCierre", "Flag Cierre a Ciegas", IIf(Me.chkCierre.Value, "Verdadero", "Falso"), "lPrinter", "Flag Kitchen Printer", IIf(Me.chkPrinter.Value, "Verdadero", "Falso"), "lAlmacen", "Flag Enlace Almacen", IIf(chkAlmacen.Value, "Verdadero", "Falso"), "lAlertaVenta", "Flag Alerta Venta", IIf(chkAlertaVenta.Value, "Verdadero", "Falso"), "tCabPreCuenta", "Texto para Cabecera Unica Precuenta", Me.txtCabPreCuenta, "lInfhotel", "Flag Enlace Infhotel", IIf(chkInfhotel.Value, "Verdadero", "Falso"), _
                "lequivalencia", "Flag Muestra Equivalencia", IIf(ChkEquivalencia.Value, "Verdadero", "Falso"), "lComboGeneral", "Flag Listado General Combos", IIf(chkComboGeneral.Value, "Verdadero", "Falso"), "lMultiLocal", "Flag Multi Local", IIf(chkMultiLocal.Value, "Verdadero", "Falso"), "lClub", "Flag Club", IIf(chkClub.Value, "Verdadero", "Falso"), "tContribuyenteEspecial", "Contribuyente Especial", txtContribuyenteEspecial.Text, "fContribuyenteEspecial", "Fecha Contribuyente Especial", dtpContribuyenteEspecial.Value, "lMobileUnidadNegocio", "Mobile Filtro por Unidad Negocio", IIf(Me.chkMUnidadNegocio.Value, "Verdadero", "Falso"), "lMobilePasswordCCaja", "Mobile Contrasenia Cambio Caja", IIf(Me.chkMCCaja.Value, "Verdadero", "Falso"), "lActivaConsultaDescargo", "Consulta Descargo al Cierre Turno", IIf(Me.chkConsultaDescargo.Value, "Verdadero", "Falso"), _
                "nCabeceraV", "Impresion Factura Variable Cabecera", txtCabeceraV.Text, "nItemV", "Impresion Factura Variable Items", txtItemV.Text, "nPieV", "Impresion Factura Variable Pie", txtPieV.Text, "lFacturacionE", "Flag Facturacion Electronica", IIf(Me.chkFacturacionE.Value, "Verdadero", "Falso"), "lControlUsuario", "Flag Control Usuario", IIf(Me.chkControlUsuario.Value, "Verdadero", "Falso"), "lHoraEntregaDelivery", "Flag Hora Entrega", IIf(Me.chkHoraEntrega.Value, "Verdadero", "Falso"), _
                "lHuellaDigital", "Flag Digital Persona", IIf(Me.chkDigital.Value, "Verdadero", "Falso"), "lHuellaSecugen", "Flag Secugen", IIf(Me.chkSecugen.Value, "Verdadero", "Falso"), "tAgenteRetencion", "Texto Agente Retencion", txtRetencion.Text, "lCheffFiltroSalon", "Cheff-Flag Filtro Salon", IIf(Me.chkCheffFiltroSalon.Value, "Verdadero", "Falso"), "lCheffFiltroSubGrupo", "Cheff-Flag Filtro Sub-Grupo", IIf(Me.chkCheffFiltroSubGrupo.Value, "Verdadero", "Falso"), "lFESpring", "Cheff-Flag Facturación Electronica Spring", IIf(Me.chkFESpring.Value, "Verdadero", "Falso"), "lFECarbajal", "Cheff-Flag Facturación Electronica Carbajal", IIf(Me.chkFECarbajal.Value, "Verdadero", "Falso"), "lFETCI", "Facturación Electronica TCI", IIf(Me.chkfeTCI.Value, "Verdadero", "Falso"), _
                "lDesactivaNCFP", "Desactiva NC forma de Pago", IIf(Me.chkDesNCPG.Value, "Verdadero", "Falso"), "lFEBiz", "Facturacion Electronica Bizlinks", IIf(Me.chkFEBiz.Value, "Verdadero", "Falso"), "tCodAnticipo", "Codigo de Producto Anticipo", Me.txtCodigoAnticipo.Text, "lActivaAnticipo", "Activa Anticipo", IIf(Me.chkActivaAnticipo.Value, "Verdadero", "Falso"), "lFEGood", "Activa FE Good Hope", IIf(Me.chkFEGood.Value, "Verdadero", "Falso"), "tMaxMotorizado", "Maximo de Motorizado", Me.txtMaxMotorizado.Text, "lStockDescargo", "Valida Stock en Descargo", IIf(Me.chkValidaStock.Value, "Verdadero", "Falso"), "lFEubl21", "Activa UBL 2.1 FE paperlees", IIf(Me.chkFEubl21.Value, "Verdadero", "Falso"), "lBloqInafecto", "Bloquear inafecto", IIf(Me.chkBloqInafecto.Value, "Verdadero", "Falso"), _
                "lEstupendoFE", "Facturacion Estupendo Ecuador", IIf(Me.chkFEEstupendo.Value, "Verdadero", "Falso"), "LFEGesa", "Facturacion electronica GESA", IIf(Me.chkFEGesa.Value, "Verdadero", "Falso"), "tLey1", "Ley Aplicale 1", Me.txtLey1.Text, "tValorLey1", "Valor ley 1", Me.txtValorLey1.Text, "tRutaMotor", "Ruta de motor de integraciones", Trim(Me.txtRutaMotorIntegraciones.Text), _
                "tGlosaImprInafecto", "Glosa Inafecto", Trim(Me.txttGlosaImprInafecto.Text), "tGlosaImprExonerado", "Glosa Exonerado", Trim(Me.txttGlosaImprExonerado.Text), "nPorcentajePropina", "Porcentaje de Propina Sugerida", Trim(Me.txtPorcentajePropinaSugerida.Text), _
                "lRetencion", "Porcentaje Retencion", Trim(Me.TxtlRetencion.Text), "tMontoRetencion", "Monto de retenciones", Trim(Me.TxtMontoMinRete.Text), _
                "lCambioContrasena", "Permiso para Cambio de Contraseña", IIf(Me.ChkCambiarContrasena.Value, "Verdadero", "Falso"), _
                "lReservaPedido", "Convertir Reserva a Pedido", IIf(Me.ChlReservaPedido.Value, "Verdadero", "Falso"), "tCodTarjeta", "Codigo de Tarjeta", txtCodigoTarjeta.Text, "lMontoMaximo", "Monto Maximo de Recarga Tarjeta", TxtMontoMaximo.Text, "lNumdiasAnulacion", "Maximo de dias para anular un documento", txtNumdiasAnulacion.Text, "lBloquearNCsinEstado", "Bloquear crear NC,sin estado ACEPTADO", IIf(Me.ChklBloquearNCsinEstado.Value, "Verdadero", "Falso"), "lImprimirFormaPago", "Imprimir Forma de Pago en Bol / Fac", IIf(Me.ChklImprimirFormaPago.Value, "Verdadero", "Falso"), "lFETusFacturasAPP", "FE Argentina TusFacturasAPP", IIf(Me.chkFETusFacturasAPP.Value, "Verdadero", "Falso"), "tPosTCI", "COdigo POS TCI", Me.txtPOSTCI.Text, "lFEBizOffLine", "Bizlinks OffLine", IIf(Me.chklFEBizOffLine.Value, "Verdadero", "Falso"), _
                "lActivarIncrementoPrecios", "Activar Incremento Precios", IIf(Me.chklActivarIncrementoPrecios.Value, "Verdadero", "Falso"), "lDetraccion", "Activar Detraccion", IIf(Me.ChklDetraccion.Value, "Verdadero", "Falso"), "lMontoDetraccion", "Monto Detraccion", Trim(Me.TxtlMontoDetraccion.Text), "lPorcentajeDetraccion", "Porcentaje Detraccion", Trim(Me.TxtlPorcentajeDetraccion.Text), "tMsjDetraccion", "Mensaje Detraccion", Trim(Me.TxttMsjDetraccion.Text), "lBODato", "Flag Solicita Dato Boleta", IIf(Me.chkSolicitarDato.Value, "Verdadero", "Falso"), "lConsultaCierreInv", "Consulta Cierre Inventario Abierto antes de procesar", IIf(Me.chkControlInventario.Value, "Verdadero", "Falso"), "lActivaReciboSaldoQr", "Activa QR en recibo ingreso y saldo a favor", IIf(Me.chkReciboSaldoFavor.Value, "Verdadero", "Falso"), "lCabeceraFacturacion", "Activa visualizacion cabecera FE Ecuador", IIf(Me.chkCabeceraFacturacion.Value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
      Isql = "Update TPARAMETRO Set " & _
              "tRazonSocial = '" & txtSocial.Text & "',  tRazonComercial ='" & txtComercial.Text & "',  tDireccion ='" & txtDireccion.Text & "', tDireccion2 ='" & txtDireccion2.Text & "', " & _
              "tIdentificacionTributaria ='" & sPasa & "',  tMonedaN ='" & txtMonedaN.Text & "',  lActivaConsultaDescargo =" & IIf(chkConsultaDescargo.Value, 1, 0) & " ,  " & _
              "tMonN ='" & txtMonN.Text & "',  tMonedaE ='" & txtMonedaE.Text & "', tMonE ='" & txtMonE.Text & "',  tMonedaE2 ='" & txtMonedaE2.Text & "', tMonE2 ='" & txtMonE2.Text & "',  tMonedaE3 ='" & txtMonedaE3.Text & "', tMonE3 ='" & txtMonE3.Text & "', " & _
              "tTelefono ='" & txtTelefono.Text & "', tEmail ='" & txtEmail.Text & "', tWebPage ='" & txtWebPage.Text & "', " & _
              "tPie ='" & txtPie.Text & "',  tPieprecuenta ='" & txtPiePreCuenta.Text & "', " & _
              "lemailconfirmacion=" & IIf(Me.chkConfirmacion.Value, 1, 0) & ", lemailrecordatorio=" & IIf(Me.chkRecordatorio.Value, 1, 0) & ", lemailagradecimiento=" & IIf(Me.chkAgradecimiento.Value, 1, 0) & " , " & _
              "temailconfirmacion='" & Trim(Me.txtConfirmacion.Text) & "', temailrecordatorio='" & Trim(Me.txtRecordatorio.Text) & "', temailagradecimiento='" & Trim(Me.txtAgradecimiento.Text) & "' , " & _
              "tImpuesto1 ='" & txtDImp1.Text & "', tImpuesto2 ='" & txtDImp2.Text & "', tImpuesto3 ='" & txtDImp3.Text & "', " & _
              "nDelivery =" & val(txtDelivery.Text) & ", nLlevar =" & val(txtllevar.Text) & ", nCanal4 =" & val(txtCanal4.Text) & ", nCanal5 =" & val(txtCanal5.Text) & ", " & _
              "nTiempoToleranciaAnf=" & val(Me.txtToleranciaReserva.Text) & ",nCorrelativo =" & val(txtCorrelativo.Text) & ",lMobileUnidadNegocio =" & IIf(Me.chkMUnidadNegocio.Value, 1, 0) & ", lMobilePasswordCCaja =" & IIf(Me.chkMCCaja.Value, 1, 0) & ",  " & _
              "Impuesto1 =" & val(txtIImp1.Text) & ", Impuesto2 =" & val(txtIImp2.Text) & ", Impuesto3 =" & val(txtIImp3.Text) & ", " & _
              "tElimina ='" & txtElimina.Text & "', nAsignacionMotorizado=" & val(Me.txtAsignacionMotorizado.Text) & "," & _
              "lCierre =" & IIf(chkCierre.Value, 1, 0) & ", lequivalencia =" & IIf(ChkEquivalencia.Value, 1, 0) & ", tAgenteRetencion='" & txtRetencion & "', " & _
              "lAlmacen =" & IIf(chkAlmacen.Value, 1, 0) & ", lAlertaVenta =" & IIf(chkAlertaVenta.Value, 1, 0) & " , tCabPreCuenta ='" & txtCabPreCuenta & "', nItemGuia =" & txtItemGuia.Text & ", nCabeceraGuia =" & txtCabeceraGuia.Text & ", nDetalleGuia =" & txtDetalleGuia.Text & ", " & _
              "nLongitud = " & val(txtLongitud.Text) & ", lLongitud=" & IIf(opcLongitud(0).Value, 1, 0) & ", tContribuyenteEspecial='" & txtContribuyenteEspecial.Text & "'," & _
              "lPrinter =" & IIf(chkPrinter.Value, 1, 0) & ", lInfhotel =" & IIf(chkInfhotel.Value, 1, 0) & ",lhuelladigital=" & IIf(chkDigital.Value, 1, 0) & ",lhuellasecugen=" & IIf(chkSecugen.Value, 1, 0) & ", " & _
              "lMultiLocal =" & IIf(chkMultiLocal.Value, 1, 0) & ", lDiaContableAutomatico=" & IIf(optDCAutomatico.Value, 1, 0) & " , lDiaContableManual=" & IIf(optDCManual.Value, 1, 0) & " ,thoracierrediacontable='" & Format(Me.dtpHoraDC.Value, "HH:mm") & "', " & _
              "nItem =" & txtItem.Text & ", nCabecera =" & txtCabecera.Text & ", nDetalle =" & txtDetalle.Text & ", tPassword ='', tClub ='" & txtClub.Text & "', nPunto=" & val(txtPunto.Text) & ", nDias= " & val(txtDia.Text) & ", nDiasDelivery= " & val(txtDiaDelivery.Text) & ", lComboGeneral=" & IIf(chkComboGeneral.Value, 1, 0) & ", nTiempoMinutoCD =" & txtTiempoDelivery.Text & ", lClub = '" & chkClub.Value & "', lImprimeDiacontable='" & chkImprimeDiaContable.Value & "', fContribuyenteEspecial= '" & Format(dtpContribuyenteEspecial.Value, "yyyy/MM/dd") & "', " & _
              "nCabeceraV =" & txtCabeceraV.Text & ", nItemV =" & txtItemV.Text & ", nPieV =" & txtPieV.Text & ", lFacturacionE='" & chkFacturacionE.Value & "', lControlUsuario='" & chkControlUsuario.Value & "', lHoraEntregaDelivery='" & chkHoraEntrega.Value & "', " & _
              "tCodigoFE='" & txtCodigoFE.Text & "',tPieDocumento1='" & txtPieFE.Text & "',lAmbienteFE='" & chkAmbienteFE.Value & "',tFax='" & txtFax.Text & "', lImprimeCodigoBarras = " & xImprimeCodigoBarras & ",lEnvioAutomatico = '" & chkEnvioAutomatico.Value & "', lControlEnviosProduccion = '" & chkControlEnviosProduccion.Value & "', lActivaTarjeta = '" & chkTarjeta.Value & "', lEventos = '" & chkEventos.Value & "',lFEOfisis = '" & chkFEOfisis.Value & "',tCodigoEmpresa='" & txtCodigoEmpresa.Text & "',tCodigoTienda='" & txtCodigoTienda.Text & "',tCodigoMarca='" & txtCodigoMarca.Text & "'," & _
              "tCodigoUbigeo='" & txtCodigoUbigeo.Text & "', lPagoAntesImpresion='" & chkPagoAntesImpresion.Value & "',tCajaMesa247='" & txtCajaMesa247.Text & "',tServidorFE='" & txtServidorFE.Text & "',tBDFE='" & txtBDFE.Text & "',tAdicionMesa247='" & txtAdicionMesa247.Text & "',linteSAP ='" & ChkSAP.Value & "',tservidorSAP ='" & txtServidorSAP.Text & "', tBDSAP = '" & TxtBaseSAP.Text & "', tCodAlmcSAP = '" & TxtCodAlmcSAP.Text & "', lActivaCuenCorrienteAut= '" & ChkActCuentaCorriente.Value & "' , tCuentaContableCort = '" & TxtCuentaContable.Text & "', lVisor8 = '" & chkVisor8.Value & "', lVisortactil = '" & chktactil.Value & "', lGlosaTransGratuita='" & txtGlosaImpresion.Text & "', lvisortiempo='" & txtvisortiempo.Text & "'," & _
              "lActivaCover=" & IIf(chkCover.Value, 1, 0) & " , tMontoMinCover='" & IIf(Trim(txtMontoMinCover.Text) = "", 0, Trim(txtMontoMinCover.Text)) & "', tCodItemCover= '" & IIf(Trim(txtCodigoItemCover.Text) = "", "", Trim(txtCodigoItemCover.Text)) & "', lNcOfisis=" & IIf(chkInNC.Value, 1, 0) & ", tRutaFE='" & Trim(txtRutaImgFE.Text) & "', lCodigoQrFE = " & IIf(optOpcion(2).Value, 1, 0) & " , lactivaFechaNC= " & IIf(chkNCFecha.Value, 1, 0) & ", lCheffFiltroSalon=" & IIf(chkCheffFiltroSalon.Value, 1, 0) & ", lCheffFiltroSubGrupo=" & IIf(chkCheffFiltroSubGrupo.Value, 1, 0) & ", lFEpape=" & IIf(chkFEpape.Value, 1, 0) & ", lAnula=" & IIf(chkAnulacionNC.Value, 1, 0) & ", lDesPagoCheque= " & IIf(chkPagoCheque.Value, 1, 0) & ", lDesPagoOtro=" & IIf(chkPagoOtra.Value, 1, 0) & ", lFESpring=" & IIf(chkFESpring.Value, 1, 0) & ", tUsuarioFE='" & txtUsuarioFE.Text & "', tClaveFE='" & txtClaveFE.Text & "', lFECarbajal=" & IIf(chkFECarbajal.Value, 1, 0) & "," & "paramCarvajal = '" & Me.txtParamCarv.Text & "' " & _
              ", tCarvajalCorreos='" & Me.txtCarvajalCorreos.Text & "'" & ", lParcialNC = " & IIf(Me.chkNCParcial.Value, 1, 0) & ", lNCElimina = " & IIf(Me.chkNCElimina.Value, 1, 0) & _
              ", lValidaDNI = " & IIf(Me.chkValidaDNI.Value, 1, 0) & ", lVerTCImp = " & IIf(Me.chkTCenImp.Value, 1, 0) & ", lFETCI= " & IIf(Me.chkfeTCI.Value, 1, 0) & ", lDesactivaNCFP=" & IIf(Me.chkDesNCPG.Value, 1, 0) & ", lFEBiz=" & IIf(Me.chkFEBiz.Value, 1, 0) & ", tCodAnticipo='" & Trim(Me.txtCodigoAnticipo.Text) & "', lActivaAnticipo=" & IIf(Me.chkActivaAnticipo.Value, 1, 0) & ", lFEGood= " & IIf(Me.chkFEGood.Value, 1, 0) & ", tMaxMotorizado='" & Me.txtMaxMotorizado.Text & "', lStockDescargo=" & IIf(Me.chkValidaStock.Value, 1, 0) & " , lFEubl21=" & IIf(Me.chkFEubl21.Value, 1, 0) & " , lBloqInafecto=" & IIf(Me.chkBloqInafecto.Value, 1, 0) & " , lEstupendoFE=" & IIf(Me.chkFEEstupendo.Value, 1, 0) & " , lFEGesa=" & IIf(Me.chkFEGesa.Value, 1, 0) & ", lDescargoInsumo=" & IIf(Me.chkDescargoInsumo.Value, 1, 0) & ", tLey1='" & Trim(Me.txtLey1.Text) & "', tValorLey1='" & Trim(Me.txtValorLey1.Text) & "' , tCorreoPape ='" & Trim(Me.TxtCorreoPape.Text) & "', tCodTarjeta ='" & Trim(Me.txtCodigoTarjeta.Text) & "'"
              'parametro carvajal arriba

      Cn.Execute Isql
      
      ' Nuevos cambios
      Isql = "Update TPARAMETRO Set tRutaMotor='" & Trim(Me.txtRutaMotorIntegraciones.Text) & "', trutalogo='" & Trim(Me.TxtRutaLogo.Text) & "', trutaimadia='" & Trim(Me.TxtRutaImaDia.Text) & "' " & _
                ", tUrlApiWeb='" & Trim(Me.txtUrlApi.Text) & "', tCodigoTiendaApiWeb='" & Trim(Me.txtCodigoTiendaApi.Text) & "', lActivaCostoEnvioPatio=" & Me.ChkItemDelivery.Value & ", tCodigoItemCostoEnviopatio='" & Me.txtItemDelivery.Text & "' " & _
                ", tGlosaImprInafecto ='" & Trim(Me.txttGlosaImprInafecto.Text) & "' , tGlosaImprExonerado ='" & Trim(Me.txttGlosaImprExonerado.Text) & "'   , nPorcentajePropina='" & Me.txtPorcentajePropinaSugerida.Text & "',  lRetencion= '" & Trim(TxtlRetencion.Text) & "',  tMontoRetencion= '" & Trim(TxtMontoMinRete.Text) & "'  " & _
                ", lReservaPedido= '" & Me.ChlReservaPedido.Value & "', lMontoMaximo= '" & Me.TxtMontoMaximo.Text & "', lNumdiasAnulacion= '" & Me.txtNumdiasAnulacion.Text & "', lBloquearNCsinEstado= '" & Me.ChklBloquearNCsinEstado.Value & "', lImprimirFormaPago= '" & Me.ChklImprimirFormaPago.Value & "'" & _
                ", lFETusFacturasAPP= '" & Me.chkFETusFacturasAPP.Value & "', lActivaMotivoAnulacion= '" & Me.ChklActivaMotivoAnulacion.Value & "', lFEBizOffLine= '" & Me.chklFEBizOffLine.Value & "', lActivarIncrementoPrecios = '" & Me.chklActivarIncrementoPrecios.Value & "', tPosTCI= '" & Me.txtPOSTCI.Text & "'" & _
                ", lDetraccion = " & IIf(Me.ChklDetraccion.Value, 1, 0) & ", lMontoDetraccion=" & Trim(Me.TxtlMontoDetraccion.Text) & ", lPorcentajeDetraccion= " & Trim(Me.TxtlPorcentajeDetraccion.Text) & ", tMsjDetraccion = '" & Trim(Me.TxttMsjDetraccion.Text) & "', lBODato = '" & Me.chkSolicitarDato.Value & "',lCabeceraFacturacion = '" & Me.chkCabeceraFacturacion.Value & "'" & _
                ", lCambioContrasena = " & IIf(Me.ChkCambiarContrasena.Value, 1, 0) & ", lvisibleimpuesto1 = " & Me.chkVisibleimpuesto(1).Value & ", lvisibleimpuesto2 = " & Me.chkVisibleimpuesto(2).Value & ", lvisibleimpuesto3 = " & Me.chkVisibleimpuesto(3).Value & ", lConsultaCierreInv = " & Me.chkControlInventario.Value & "" & _
                ", lActivaReciboSaldoQr = " & IIf(Me.chkReciboSaldoFavor.Value, 1, 0) & ", tUrlImgInforestApp = '" & Me.tUrlImgInforestApp.Text & "',lActivaTimer= '" & Me.chkTimer.Value & "', tTiempoTimer= '" & Me.txtTimer.Text & "'"
      Cn.Execute Isql

      
      Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonN.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '01'"
      Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonE.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '02'"
      
      If pais = "003" Then
        Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonE2.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '03'"
        Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonE3.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '04'"
      End If

      sRazonSocial = txtSocial.Text
      sRazonComercial = txtComercial.Text
      sDireccion = txtDireccion.Text
      sDireccion2 = txtDireccion2.Text
      sRUC = txtRUC.Text
      sMonN = txtMonN.Text
      sMonedaN = txtMonedaN.Text
      sMonE = txtMonE.Text
      sMonedaE = txtMonedaE.Text
      sMonE2 = txtMonE2.Text
      sMonedaE2 = txtMonedaE2.Text
      sMonE3 = txtMonE3.Text
      sMonedaE3 = txtMonedaE3.Text
      sImpuesto1 = txtDImp1.Text
      sImpuesto2 = txtDImp2.Text
      sImpuesto3 = txtDImp3.Text
      nPorcentaje1 = val(txtIImp1.Text)
      nPorcentaje2 = val(txtIImp2.Text)
      nPorcentaje3 = val(txtIImp3.Text)
      nDELIVERY = val(txtDelivery.Text)
      nLlevar = val(txtllevar.Text)
      sElimina = txtElimina.Text
      lAlmacen = chkAlmacen.Value
      lAlertaVenta = chkAlertaVenta.Value
      lInfhotel = chkInfhotel.Value
      lPrinter = chkPrinter.Value
      
      '---SAP----
      lSAP = ChkSAP.Value
      sServidorSAp = Me.txtServidorSAP.Text
      sBdSAP = Me.TxtBaseSAP.Text
      sCodSap = Me.TxtCodAlmcSAP.Text
      '-----visor-----
      
      lvisor = chkVisor8.Value
      
      '------
      'huella
      
      lActivaConsultaDescargo = chkConsultaDescargo.Value
      lHuellaDigitalPersona = chkDigital.Value
      lHuellaSecugen = chkSecugen.Value
      
      
      'FACTURACION ELECTRONICA
      lFacturacionE = chkFacturacionE.Value
      tCodigoFE = txtCodigoFE.Text
      tPieDocumento1 = txtPieFE.Text
      lAmbienteProduccion = chkAmbienteFE.Value
      
      sLey1 = Trim(Me.txtLey1.Text)
      sValorLey1 = Trim(Me.txtValorLey1.Text)
      
      Screen.MousePointer = vbDefault
      MsgBox "Parámetros Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub cmdValidar_Click(Index As Integer)
If txtServidorFE = "" Then MsgBox "Ingrese el Nombre del Servidor...", vbExclamation, sMensaje: txtServidorFE.SetFocus: Exit Sub
If txtBDFE = "" Then MsgBox "Ingrese el Nombre de la Base de Datos...", vbExclamation, sMensaje: txtBDFE.SetFocus: Exit Sub
If txtUsuarioFE = "" Then MsgBox "Ingrese el Nombre del Usuario Sql...", vbExclamation, sMensaje: txtUsuarioFE.SetFocus: Exit Sub
If txtClaveFE = "" Then MsgBox "Ingrese el Password del Usuario Sql...", vbExclamation, sMensaje: txtClaveFE.SetFocus: Exit Sub
If (validaConexionSistemaExterno(txtServidorFE, txtBDFE, txtUsuarioFE, txtClaveFE)) = False Then
    MsgBox "No se puede establecer conexión con: " & txtServidorFE, vbCritical, sMensaje
Else
    MsgBox "Prueba de conexón satisfactoria con el Servidor " & txtServidorFE, vbInformation, sMensaje
End If
End Sub

Private Sub Form_Load()
Centrar Me
'parametro carvajal
Me.fraPaCarvajal.Visible = False

Frame(5).Caption = " Configuración de " & Guia

If pais = "000" Then
 Me.FrmFacPeru.Visible = True
End If
If pais = "002" Then
 Me.FrmFacEcuador.Visible = True
 Me.FrmCamposFacturacion.Visible = True
End If

'FE ARGENTINA GCAA 14/03/2022
If pais = "003" Then
 Me.FrmFacArgentina.Visible = True
End If
 
Isql = "select * from TPARAMETRO"
Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   
txtComercial.Text = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
txtSocial.Text = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
txtDireccion.Text = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
txtDireccion2.Text = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
txtRUC.Text = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
txtMonN.Text = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
txtMonedaN.Text = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
txtMonE.Text = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
txtMonedaE.Text = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
txtMonE2.Text = IIf(IsNull(RsParametro!tMonE2), "", RsParametro!tMonE2)
txtMonedaE2.Text = IIf(IsNull(RsParametro!tMonedaE2), "", RsParametro!tMonedaE2)
txtMonE3.Text = IIf(IsNull(RsParametro!tMonE3), "", RsParametro!tMonE3)
txtMonedaE3.Text = IIf(IsNull(RsParametro!tMonedaE3), "", RsParametro!tMonedaE3)
txtTelefono.Text = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
txtEmail.Text = IIf(IsNull(RsParametro!temail), "", RsParametro!temail)
txtWebPage.Text = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
txtPie.Text = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
txtPiePreCuenta.Text = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
txtDImp1.Text = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
txtDImp2.Text = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
txtDImp3.Text = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
txtIImp1.Text = IIf(IsNull(RsParametro!IMPUESTO1), 0, RsParametro!IMPUESTO1)
txtIImp2.Text = IIf(IsNull(RsParametro!IMPUESTO2), 0, RsParametro!IMPUESTO2)
txtIImp3.Text = IIf(IsNull(RsParametro!IMPUESTO3), 0, RsParametro!IMPUESTO3)
txtCorrelativo.Text = IIf(IsNull(RsParametro!nCorrelativo), "", RsParametro!nCorrelativo)
txtDelivery.Text = IIf(IsNull(RsParametro!nDELIVERY), 0, RsParametro!nDELIVERY)
txtllevar.Text = IIf(IsNull(RsParametro!nLlevar), 0, RsParametro!nLlevar)
txtCanal4.Text = IIf(IsNull(RsParametro!nCanal4), 0, RsParametro!nCanal4)
txtCanal5.Text = IIf(IsNull(RsParametro!nCanal5), 0, RsParametro!nCanal5)
txtElimina.Text = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
txtItem.Text = IIf(IsNull(RsParametro!nItem), 0, RsParametro!nItem)
txtTimer.Text = IIf(IsNull(RsParametro!tTiempoTimer), 1, RsParametro!tTiempoTimer)
chkPrinter.Value = IIf(IsNull(RsParametro!lPrinter), 0, IIf(RsParametro!lPrinter = True, 1, 0))
txtLongitud.Text = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
chkAlmacen.Value = IIf(IsNull(RsParametro!lAlmacen), 0, IIf(RsParametro!lAlmacen = True, 1, 0))
chkAlertaVenta.Value = IIf(IsNull(RsParametro!lAlertaVenta), 0, IIf(RsParametro!lAlertaVenta = True, 1, 0))
chkInfhotel.Value = IIf(IsNull(RsParametro!lInfhotel), 0, IIf(RsParametro!lInfhotel = True, 1, 0))
chkControlInventario.Value = IIf(IsNull(RsParametro!lConsultaCierreInv), 0, IIf(RsParametro!lConsultaCierreInv = True, 1, 0))
chkMultiLocal.Value = IIf(IsNull(RsParametro!lmultilocal), 0, IIf(RsParametro!lmultilocal = True, 1, 0))
ChkCambiarContrasena = IIf(IsNull(RsParametro!lCambioContrasena), 0, IIf(RsParametro!lCambioContrasena = True, 1, 0))
chkCierre = IIf(IsNull(RsParametro!lCierre), 0, IIf(RsParametro!lCierre = True, 1, 0))
chkSolicitarDato = IIf(IsNull(RsParametro!lBODato), 0, IIf(RsParametro!lBODato = True, 1, 0))
'Campo para cabecera documento
chkCabeceraFacturacion.Value = IIf(IsNull(RsParametro!lCabeceraFacturacion), 0, IIf(RsParametro!lCabeceraFacturacion = True, 1, 0))

txtClub.Text = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
txtPunto.Text = Format(IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto), "#,##0.00")
txtDia.Text = Format(IIf(IsNull(RsParametro!nDias), 1, RsParametro!nDias), "#,##0")
txtTiempoDelivery.Text = Format(IIf(IsNull(RsParametro!nTiempoMinutoCD), 0, RsParametro!nTiempoMinutoCD), "##0")
txtDiaDelivery.Text = Format(IIf(IsNull(RsParametro!nDiasDelivery), 1, RsParametro!nDiasDelivery), "#,##0")
txtCabPreCuenta.Text = IIf(IsNull(RsParametro!tCabPreCuenta), "", RsParametro!tCabPreCuenta)
'parametro cravajal
Me.txtParamCarv = IIf(IsNull(RsParametro!paramCarvajal), "", RsParametro!paramCarvajal)

txtFax.Text = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
' glosa de impresion transferencia gratuita
Me.txtGlosaImpresion.Text = IIf(IsNull(RsParametro!lGlosaTransGratuita), "", RsParametro!lGlosaTransGratuita)

'-------------
Me.txtRetencion.Text = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)

txtCabecera.Text = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
txtDetalle.Text = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
txtItemGuia.Text = IIf(IsNull(RsParametro!nItemGuia), 0, RsParametro!nItemGuia)
txtCabeceraGuia.Text = IIf(IsNull(RsParametro!nCabeceraGuia), 0, RsParametro!nCabeceraGuia)
txtDetalleGuia.Text = IIf(IsNull(RsParametro!nDetalleGuia), 0, RsParametro!nDetalleGuia)
ChkEquivalencia.Value = IIf(IsNull(RsParametro!lEquivalencia), 0, IIf(RsParametro!lEquivalencia = True, 1, 0))
chkComboGeneral.Value = IIf(IsNull(RsParametro!lComboGeneral), 0, IIf(RsParametro!lComboGeneral = True, 1, 0))
txtContribuyenteEspecial = IIf(IsNull(RsParametro!tContribuyenteEspecial), 0, (RsParametro!tContribuyenteEspecial))
dtpContribuyenteEspecial = IIf(IsNull(RsParametro!fContribuyenteEspecial), 2, (RsParametro!fContribuyenteEspecial))

Me.chkMUnidadNegocio.Value = IIf(IsNull(RsParametro!lMobileUnidadNegocio), 0, IIf(RsParametro!lMobileUnidadNegocio = True, 1, 0))
Me.chkMCCaja.Value = IIf(IsNull(RsParametro!lMobilePasswordCCaja), 0, IIf(RsParametro!lMobilePasswordCCaja = True, 1, 0))

chkConsultaDescargo.Value = IIf(IsNull(RsParametro!lActivaConsultaDescargo), 0, IIf(RsParametro!lActivaConsultaDescargo = True, 1, 0))
'huella
chkDigital.Value = IIf(IsNull(RsParametro!lHUELLADIGITAL), 0, IIf(RsParametro!lHUELLADIGITAL = True, 1, 0))
chkSecugen.Value = IIf(IsNull(RsParametro!lHuellaSecugen), 0, IIf(RsParametro!lHuellaSecugen = True, 1, 0))
Me.txtCarvajalCorreos.Text = IIf(IsNull(RsParametro!tCarvajalCorreos), "", RsParametro!tCarvajalCorreos)

chkValidaDNI.Value = IIf(IsNull(RsParametro!lValidaDNI), 0, IIf(RsParametro!lValidaDNI = True, 1, 0))
Me.chkTCenImp.Value = IIf(IsNull(RsParametro!lVerTCImp), 0, IIf(RsParametro!lVerTCImp = True, 1, 0))

Me.chkFEubl21.Value = IIf(IsNull(RsParametro!lFEubl21), 0, IIf(RsParametro!lFEubl21 = True, 1, 0))

      
                                                                                     

If IsNull(RsParametro!lLongitud) Or RsParametro!lLongitud = 0 Then
  opcLongitud(0).Value = 0
  opcLongitud(1).Value = 1
Else
  opcLongitud(0).Value = 1
  opcLongitud(1).Value = 0
End If

Me.optDCAutomatico.Value = IIf(IsNull(RsParametro!lDiaContableAutomatico), 0, RsParametro!lDiaContableAutomatico)
Me.optDCManual.Value = IIf(IsNull(RsParametro!lDiaContablemanual), 0, RsParametro!lDiaContablemanual)
Me.dtpHoraDC.Value = IIf(IsNull(RsParametro!tHoraCierreDiaContable), "00:00", RsParametro!tHoraCierreDiaContable)
chkImprimeDiaContable.Value = IIf(IsNull(RsParametro!lImprimeDiaContable), 0, IIf(RsParametro!lImprimeDiaContable = True, 1, 0))

'Club
Me.chkClub.Value = IIf(IsNull(RsParametro!lClub), 0, IIf(RsParametro!lClub = True, 1, 0))

'motorizados
txtAsignacionMotorizado.Text = Format(IIf(IsNull(RsParametro!nAsignacionMotorizado), 0, RsParametro!nAsignacionMotorizado), "###,##0.00")

'Formnato Variabe
txtCabeceraV.Text = IIf(IsNull(RsParametro!nCabeceraV), 0, RsParametro!nCabeceraV)
txtItemV.Text = IIf(IsNull(RsParametro!nItemV), 0, RsParametro!nItemV)
txtPieV.Text = IIf(IsNull(RsParametro!nPieV), 0, RsParametro!nPieV)

'FACTURACION ELECTRONICA
Me.chkFacturacionE.Value = IIf(IsNull(RsParametro!lFacturacionE), 0, IIf(RsParametro!lFacturacionE = True, 1, 0))
Me.txtCodigoFE.Text = IIf(IsNull(RsParametro!tCodigoFE), "000", RsParametro!tCodigoFE)
Me.txtPieFE.Text = IIf(IsNull(RsParametro!tPieDocumento1), "", RsParametro!tPieDocumento1)
Me.chkAmbienteFE.Value = IIf(IsNull(RsParametro!lAmbienteFE), 0, IIf(RsParametro!lAmbienteFE = True, 1, 0))

'Control Usuarios
Me.chkControlUsuario.Value = IIf(IsNull(RsParametro!lControlUsuario), 0, IIf(RsParametro!lControlUsuario = True, 1, 0))
Me.chkHoraEntrega.Value = IIf(IsNull(RsParametro!lHoraEntregaDelivery), 0, IIf(RsParametro!lHoraEntregaDelivery = True, 1, 0))

'anfitriona
Me.chkConfirmacion.Value = IIf(IsNull(RsParametro!lEmailConfirmacion), 0, IIf(RsParametro!lEmailConfirmacion = True, 1, 0))
Me.chkRecordatorio.Value = IIf(IsNull(RsParametro!lEmailRecordatorio), 0, IIf(RsParametro!lEmailRecordatorio = True, 1, 0))
Me.chkAgradecimiento.Value = IIf(IsNull(RsParametro!lEmailAgradecimiento), 0, IIf(RsParametro!lEmailAgradecimiento = True, 1, 0))

Me.txtConfirmacion.Text = IIf(IsNull(RsParametro!tEmailConfirmacion), "", RsParametro!tEmailConfirmacion)
Me.txtRecordatorio.Text = IIf(IsNull(RsParametro!tEmailRecordatorio), "", RsParametro!tEmailRecordatorio)
Me.txtAgradecimiento.Text = IIf(IsNull(RsParametro!tEmailAgradecimiento), "", RsParametro!tEmailAgradecimiento)
Me.txtToleranciaReserva.Text = Format(IIf(IsNull(RsParametro!nTiempoToleranciaAnf), 0, RsParametro!nTiempoToleranciaAnf), "##0")
 
lImprimeCodigoBarras = IIf(IsNull(RsParametro!lImprimeCodigoBarras), 0, IIf(RsParametro!lImprimeCodigoBarras = True, 1, 0))

Me.chkEnvioAutomatico.Value = IIf(IsNull(RsParametro!lEnvioAutomatico), 0, IIf(RsParametro!lEnvioAutomatico = True, 1, 0))

Me.chkControlEnviosProduccion.Value = IIf(IsNull(RsParametro!lControlEnviosProduccion), 0, IIf(RsParametro!lControlEnviosProduccion = True, 1, 0))


Me.chkTarjeta.Value = IIf(IsNull(RsParametro!lActivaTarjeta), 0, IIf(RsParametro!lActivaTarjeta = True, 1, 0))

Me.chkEventos.Value = IIf(IsNull(RsParametro!lEventos), 0, IIf(RsParametro!lEventos = True, 1, 0))

' --- activacion de cuentas corrientes Automatico
ChkActCuentaCorriente.Value = IIf(IsNull(RsParametro!lActivaCuenCorrienteAut), 0, IIf(RsParametro!lActivaCuenCorrienteAut = True, 1, 0))

Me.chkFEOfisis.Value = IIf(IsNull(RsParametro!lFEOfisis), 0, IIf(RsParametro!lFEOfisis = True, 1, 0))

Me.txtCodigoEmpresa.Text = IIf(IsNull(RsParametro!tCodigoEmpresa), "", RsParametro!tCodigoEmpresa)
Me.txtCodigoTienda.Text = IIf(IsNull(RsParametro!tCodigoTienda), "", RsParametro!tCodigoTienda)
Me.txtCodigoMarca.Text = IIf(IsNull(RsParametro!tCodigoMarca), "", RsParametro!tCodigoMarca)
Me.txtCodigoUbigeo.Text = IIf(IsNull(RsParametro!tCodigoUbigeo), "", RsParametro!tCodigoUbigeo)
Me.TxtCuentaContable.Text = IIf(IsNull(RsParametro!tCuentaContableCort), "", RsParametro!tCuentaContableCort)
   
Me.chkPagoAntesImpresion.Value = IIf(IsNull(RsParametro!lPagoAntesImpresion), 0, IIf(RsParametro!lPagoAntesImpresion = True, 1, 0))

Me.txtCajaMesa247.Text = IIf(IsNull(RsParametro!tCajaMesa247), "", RsParametro!tCajaMesa247)
Me.txtAdicionMesa247.Text = IIf(IsNull(RsParametro!tAdicionMesa247), "", RsParametro!tAdicionMesa247)
Me.txtServidorFE.Text = IIf(IsNull(RsParametro!tServidorFE), "", RsParametro!tServidorFE)
Me.txtBDFE.Text = IIf(IsNull(RsParametro!tBDFE), "", RsParametro!tBDFE)
'--- SAP
Me.ChkSAP.Value = IIf(IsNull(RsParametro!lInteSAP), 0, IIf(RsParametro!lInteSAP = True, 1, 0))
ChkSAP_Click
 If ChkSAP.Value = 1 Then
  'Me.txtServidorSAP.Text = IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
  'Me.TxtBaseSAP.Text = IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
   Me.TxtCodAlmcSAP.Text = IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
 Else
  'Me.txtServidorSAP.Text = "" 'IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
  'Me.TxtBaseSAP.Text = "" ' IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
  Me.TxtCodAlmcSAP.Text = "" 'IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
 End If
 '----visor-----

 Me.chkVisor8.Value = IIf(IsNull(RsParametro!lvisor8), 0, IIf(RsParametro!lvisor8 = True, 1, 0))
 Me.chktactil.Value = IIf(IsNull(RsParametro!lvisortactil), 0, IIf(RsParametro!lvisortactil = True, 1, 0))
 Me.txtvisortiempo.Text = IIf(IsNull(RsParametro!lvisortiempo), "", RsParametro!lvisortiempo)
'-------
'---- bar - cover ecuador----'
Me.chkCover.Value = IIf(IsNull(RsParametro!lActivaCover), 0, IIf(RsParametro!lActivaCover = True, 1, 0))
Me.txtMontoMinCover.Text = IIf(IsNull(RsParametro!tMontoMinCover), "", RsParametro!tMontoMinCover)
Me.txtCodigoItemCover.Text = IIf(IsNull(RsParametro!tCodItemCover), "", RsParametro!tCodItemCover)

Me.txtRutaImgFE.Text = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE)
'------------------------
' ofisis
Me.chkInNC.Value = IIf(IsNull(RsParametro!lNcOfisis), 0, IIf(RsParametro!lNcOfisis = True, 1, 0))

' notas de credito
Me.chkNCFecha.Value = IIf(IsNull(RsParametro!lactivaFechaNC), 0, IIf(RsParametro!lactivaFechaNC = True, 1, 0))
Me.chkNCParcial.Value = IIf(IsNull(RsParametro!lParcialNC), 0, IIf(RsParametro!lParcialNC = True, 1, 0))
Me.chkNCElimina.Value = IIf(IsNull(RsParametro!lNCElimina), 0, IIf(RsParametro!lNCElimina = True, 1, 0))

'cheff control
Me.chkCheffFiltroSalon.Value = IIf(IsNull(RsParametro!lCheffFiltroSalon), 0, IIf(RsParametro!lCheffFiltroSalon = True, 1, 0))
Me.chkCheffFiltroSubGrupo.Value = IIf(IsNull(RsParametro!lCheffFiltroSubGrupo), 0, IIf(RsParametro!lCheffFiltroSubGrupo = True, 1, 0))

'FE Paperlees
Me.chkFEpape.Value = IIf(IsNull(RsParametro!lFEpape), 0, IIf(RsParametro!lFEpape = True, 1, 0))
'anulacion de documentos por nota de credito
Me.chkAnulacionNC.Value = IIf(IsNull(RsParametro!lAnula), 0, IIf(RsParametro!lAnula = True, 1, 0))
'   If Me.chkFEpape.value = False Then
'        Me.chkAnulacionNC.Enabled = False
'
'   End If
'--------------------------------------------

Me.chkPagoCheque.Value = IIf(IsNull(RsParametro!lDesPagoCheque), 0, IIf(RsParametro!lDesPagoCheque = True, 1, 0))
Me.chkPagoOtra.Value = IIf(IsNull(RsParametro!lDesPagoOtro), 0, IIf(RsParametro!lDesPagoOtro = True, 1, 0))

'FE Spring
Me.chkFESpring.Value = IIf(IsNull(RsParametro!lFESpring), 0, IIf(RsParametro!lFESpring = True, 1, 0))
Me.txtUsuarioFE.Text = IIf(IsNull(RsParametro!tUsuarioFE), "", RsParametro!tUsuarioFE)
Me.txtClaveFE.Text = IIf(IsNull(RsParametro!tClaveFE), "", RsParametro!tClaveFE)

'FE Carbajal
Me.chkFECarbajal.Value = IIf(IsNull(RsParametro!lFECarbajal), 0, IIf(RsParametro!lFECarbajal = True, 1, 0))

'FE TCI
Me.chkfeTCI.Value = IIf(IsNull(RsParametro!lFETCI), 0, IIf(RsParametro!lFETCI = True, 1, 0))

Me.chkDesNCPG.Value = IIf(IsNull(RsParametro!lDesactivaNCFP), 0, IIf(RsParametro!lDesactivaNCFP = True, 1, 0))
Me.chkFEBiz.Value = IIf(IsNull(RsParametro!lFEBiz), 0, IIf(RsParametro!lFEBiz = True, 1, 0))
Me.txtCodigoAnticipo.Text = IIf(IsNull(RsParametro!tCodAnticipo), "", RsParametro!tCodAnticipo)
Me.chkActivaAnticipo.Value = IIf(IsNull(RsParametro!lActivaAnticipo), 0, IIf(RsParametro!lActivaAnticipo = True, 1, 0))
Me.chkFEGood.Value = IIf(IsNull(RsParametro!lFEGood), 0, IIf(RsParametro!lFEGood = True, 1, 0))
Me.txtMaxMotorizado.Text = IIf(IsNull(RsParametro!tMaxMotorizado), "", RsParametro!tMaxMotorizado)
Me.chkValidaStock.Value = IIf(IsNull(RsParametro!lStockDescargo), 0, IIf(RsParametro!lStockDescargo = True, 1, 0))
Me.chkBloqInafecto.Value = IIf(IsNull(RsParametro!lBloqInafecto), 0, IIf(RsParametro!lBloqInafecto = True, 1, 0))
Me.chkFEEstupendo.Value = IIf(IsNull(RsParametro!lEstupendoFE), 0, IIf(RsParametro!lEstupendoFE = True, 1, 0))
Me.chkFEGesa.Value = IIf(IsNull(RsParametro!lFEGesa), 0, IIf(RsParametro!lFEGesa = True, 1, 0))
Me.chkDescargoInsumo.Value = IIf(IsNull(RsParametro!lDescargoinsumo), 0, IIf(RsParametro!lDescargoinsumo = True, 1, 0))

Me.txtLey1.Text = IIf(IsNull(RsParametro!tLey1), "", RsParametro!tLey1)
Me.txtValorLey1.Text = IIf(IsNull(RsParametro!tValorLey1), "", RsParametro!tValorLey1)
Me.txtRutaMotorIntegraciones = IIf(IsNull(RsParametro!tRutaMotor), "", RsParametro!tRutaMotor)

Me.TxtRutaLogo.Text = IIf(IsNull(RsParametro!trutalogo), "", RsParametro!trutalogo)
Me.TxtRutaImaDia.Text = IIf(IsNull(RsParametro!trutaimadia), "", RsParametro!trutaimadia)
Me.txtUrlApi.Text = IIf(IsNull(RsParametro!tUrlApiWeb), "", RsParametro!tUrlApiWeb)
Me.txtCodigoTiendaApi.Text = IIf(IsNull(RsParametro!tCodigoTiendaApiWeb), "", RsParametro!tCodigoTiendaApiWeb)
Me.txtItemDelivery.Text = IIf(IsNull(RsParametro!tCodigoItemCostoEnvioPatio), "", RsParametro!tCodigoItemCostoEnvioPatio)
Me.ChkItemDelivery.Value = IIf(IsNull(RsParametro!lActivaCostoEnvioPatio), 0, IIf(RsParametro!lActivaCostoEnvioPatio = True, 1, 0))
Me.txttGlosaImprInafecto.Text = IIf(IsNull(RsParametro!tGlosaImprInafecto), "", RsParametro!tGlosaImprInafecto)
Me.txttGlosaImprExonerado.Text = IIf(IsNull(RsParametro!tGlosaImprExonerado), "", RsParametro!tGlosaImprExonerado)
'gcaa 06012021
Me.TxtCorreoPape.Text = IIf(IsNull(RsParametro!TCorreoPape), "", RsParametro!TCorreoPape)
Me.txtPorcentajePropinaSugerida.Text = IIf(IsNull(RsParametro!nPorcentajePropina), "", RsParametro!nPorcentajePropina)
Me.TxtlRetencion.Text = IIf(IsNull(RsParametro!lRetencion), "", RsParametro!lRetencion)
Me.TxtMontoMinRete.Text = IIf(IsNull(RsParametro!tMontoRetencion), "", RsParametro!tMontoRetencion)
Me.ChlReservaPedido.Value = IIf(IsNull(RsParametro!lReservaPedido), 0, IIf(RsParametro!lReservaPedido = True, 1, 0))
Me.txtCodigoTarjeta.Text = IIf(IsNull(RsParametro!tCodTarjeta), "", RsParametro!tCodTarjeta)
Me.txtNumdiasAnulacion.Text = IIf(IsNull(RsParametro!lNumdiasAnulacion), 0, RsParametro!lNumdiasAnulacion)
Me.ChklBloquearNCsinEstado.Value = IIf(IsNull(RsParametro!lBloquearNCsinEstado), 0, IIf(RsParametro!lBloquearNCsinEstado = True, 1, 0))
Me.ChklImprimirFormaPago.Value = IIf(IsNull(RsParametro!lImprimirFormaPago), 0, IIf(RsParametro!lImprimirFormaPago = True, 1, 0))
Me.TxtMontoMaximo.Text = IIf(IsNull(RsParametro!lMontoMaximo), "", RsParametro!lMontoMaximo)
Me.chkFETusFacturasAPP.Value = IIf(IsNull(RsParametro!lFETusFacturasAPP), 0, IIf(RsParametro!lFETusFacturasAPP = True, 1, 0))
Me.ChklActivaMotivoAnulacion.Value = IIf(IsNull(RsParametro!lActivaMotivoAnulacion), 0, IIf(RsParametro!lActivaMotivoAnulacion = True, 1, 0))
Me.chkTimer.Value = IIf(IsNull(RsParametro!lActivaTimer), 0, IIf(RsParametro!lActivaTimer = True, 1, 0))
Me.chklFEBizOffLine.Value = IIf(IsNull(RsParametro!lFEBizOffLine), 0, IIf(RsParametro!lFEBizOffLine = True, 1, 0))
Me.chklActivarIncrementoPrecios.Value = IIf(IsNull(RsParametro!lActivarIncrementoPrecios), 0, IIf(RsParametro!lActivarIncrementoPrecios = True, 1, 0))

Me.ChklDetraccion.Value = IIf(IsNull(RsParametro!lDetraccion), 0, IIf(RsParametro!lDetraccion = True, 1, 0))
Me.TxtlMontoDetraccion.Text = IIf(IsNull(RsParametro!lMontoDetraccion), "0", RsParametro!lMontoDetraccion)
Me.TxtlPorcentajeDetraccion.Text = IIf(IsNull(RsParametro!lPorcentajeDetraccion), "0", RsParametro!lPorcentajeDetraccion)
Me.TxttMsjDetraccion.Text = IIf(IsNull(RsParametro!tMsjDetraccion), "0", RsParametro!tMsjDetraccion)

Me.txtPOSTCI.Text = IIf(IsNull(RsParametro!tPosTCI), "", RsParametro!tPosTCI)
Me.chkVisibleimpuesto(1).Value = IIf(IsNull(RsParametro!lvisibleimpuesto1), 0, IIf(RsParametro!lvisibleimpuesto1 = True, 1, 0))
Me.chkVisibleimpuesto(2).Value = IIf(IsNull(RsParametro!lvisibleimpuesto2), 0, IIf(RsParametro!lvisibleimpuesto2 = True, 1, 0))
Me.chkVisibleimpuesto(3).Value = IIf(IsNull(RsParametro!lvisibleimpuesto3), 0, IIf(RsParametro!lvisibleimpuesto3 = True, 1, 0))

Me.chkReciboSaldoFavor.Value = IIf(IsNull(RsParametro!lActivaReciboSaldoQr), 0, IIf(RsParametro!lActivaReciboSaldoQr = True, 1, 0))
Me.tUrlImgInforestApp.Text = IIf(IsNull(RsParametro!tUrlImgInforestApp), "", RsParametro!tUrlImgInforestApp)

   If lImprimeCodigoBarras Then
        optOpcion(0).Value = True
        optOpcion(1).Value = False
   Else
        optOpcion(0).Value = False
        optOpcion(1).Value = True
   End If
   
   If pais = "002" Or pais = "001" Then
        FrameTipoImpresion.Visible = False
   End If
   
    If IIf(IsNull(RsParametro!lCodigoQrFE), 0, IIf(RsParametro!lCodigoQrFE = True, 1, 0)) = 1 Then
        optOpcion(2).Value = True
    Else
        optOpcion(2).Value = False
    End If
    
    If Not lAlmacen Then
        'Frame35.Enabled = False
    End If
    
    SSTab1.Tab = 0
    
    If pais = "003" Then
        chkValidaDNI.Visible = False
        txtMonE2.Visible = True
        txtMonE3.Visible = True
        txtMonedaE2.Visible = True
        txtMonedaE3.Visible = True
        Label(66).Visible = True
        Label(67).Visible = True
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RsParametro.Close
   Set RsParametro = Nothing
   Set frmParametro = Nothing
End Sub





Private Sub optDCAutomatico_Click()
        If optDCAutomatico.Value = True Then
            Me.dtpHoraDC.Enabled = True
            Me.dtpHoraDC.SetFocus
        Else
            Me.dtpHoraDC.Enabled = False
        End If
End Sub

Private Sub optDCManual_Click()
        If optDCManual.Value = True Then
              Me.dtpHoraDC.Value = "00:00"
              Me.dtpHoraDC.Enabled = False
        Else
         Me.dtpHoraDC.Enabled = True
          Me.dtpHoraDC.SetFocus
            
        End If
End Sub

Private Sub txtChk_LostFocus()
  If val(txtChk.Text) < 0 Or val(txtChk.Text) > 60 Then
     MsgBox "Rango Erroneo", vbExclamation, sMensaje
     txtChk.SetFocus
  End If
End Sub


Private Sub txtAsignacionMotorizado_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtAsignacionMotorizado
End Sub

Private Sub txtMaxMotorizado_Change()
    If Not IsNumeric(Me.txtMaxMotorizado.Text) Then
        Me.txtMaxMotorizado.Text = ""
    End If
    Me.txtMaxMotorizado.Text = Trim(Me.txtMaxMotorizado.Text)
    Me.txtMaxMotorizado.SelStart = Len(Me.txtMaxMotorizado)
End Sub

Private Sub txtPunto_LostFocus()
   If val(txtPunto.Text) > 1 Then
      txtPunto.Text = Format(txtPunto.Text, "#,##0.00")
   Else
     txtPunto.Text = "1.00"
   End If
End Sub

Private Sub txtToleranciaReserva_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
    Numerico KeyAscii, txtToleranciaReserva
End Sub


Private Sub txtValorLey1_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
    Numerico KeyAscii, txtValorLey1
End Sub


