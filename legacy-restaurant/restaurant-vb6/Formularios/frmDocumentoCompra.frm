VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocumentoCompra 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Documentos"
   ClientHeight    =   8100
   ClientLeft      =   3495
   ClientTop       =   3780
   ClientWidth     =   13545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocumentoCompra.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDocumentoCompra.frx":000C
   ScaleHeight     =   8100
   ScaleWidth      =   13545
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOpcion 
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
      Height          =   520
      Index           =   9
      Left            =   12360
      MousePointer    =   99  'Custom
      Picture         =   "frmDocumentoCompra.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1020
   End
   Begin VB.Frame fraCabecera 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   13485
      Begin VB.ComboBox CmbMes 
         Height          =   330
         ItemData        =   "frmDocumentoCompra.frx":0408
         Left            =   8265
         List            =   "frmDocumentoCompra.frx":040A
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   915
         Width           =   1440
      End
      Begin VB.TextBox txtTransporte 
         BackColor       =   &H00F0EADB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7005
         MaxLength       =   30
         TabIndex        =   95
         ToolTipText     =   "Monto de Transporte"
         Top             =   915
         Width           =   600
      End
      Begin VB.TextBox txtMontoControl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   11115
         TabIndex        =   92
         Text            =   "0.00"
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Frame lblCodigoAutori 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   165
         Left            =   11670
         TabIndex        =   89
         Top             =   855
         Width           =   1725
      End
      Begin VB.Frame lblCodigoControl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   165
         Left            =   11670
         TabIndex        =   85
         Top             =   420
         Width           =   1725
      End
      Begin VB.TextBox txtFechaProceso 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   10155
         MaxLength       =   20
         TabIndex        =   84
         Top             =   510
         Width           =   1485
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   1950
         Picture         =   "frmDocumentoCompra.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ayuda de Búsqueda/Adición"
         Top             =   915
         UseMaskColor    =   -1  'True
         Width           =   380
      End
      Begin VB.TextBox TxtNombre 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2385
         TabIndex        =   7
         ToolTipText     =   "Razón Social del Proveedor"
         Top             =   540
         Width           =   5670
      End
      Begin VB.TextBox TxtProveedor 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         ToolTipText     =   "Código del Proveedor"
         Top             =   915
         Width           =   885
      End
      Begin VB.TextBox TxtFPago 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   290
         Left            =   13890
         TabIndex        =   61
         ToolTipText     =   "Fecha de Cancelación del Documento"
         Top             =   960
         Visible         =   0   'False
         Width           =   1700
      End
      Begin MSComCtl2.DTPicker DtpFCancelacion 
         Height          =   315
         Left            =   10800
         TabIndex        =   60
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54788097
         CurrentDate     =   38051
      End
      Begin MSComCtl2.DTPicker DtpEmision 
         Height          =   330
         Left            =   8805
         TabIndex        =   4
         ToolTipText     =   "Fecha de Emisión del Documento"
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16710377
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54788099
         CurrentDate     =   38051
      End
      Begin MSComCtl2.DTPicker DtpIngreso 
         Height          =   330
         Left            =   8805
         TabIndex        =   5
         ToolTipText     =   "Fecha de Ingreso al Almacen/Sub-Almacen"
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16710377
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54788099
         CurrentDate     =   38051
      End
      Begin VB.TextBox TxtGlosaDocumento 
         BackColor       =   &H00F0EADB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         MaxLength       =   150
         TabIndex        =   13
         ToolTipText     =   "Referencia/Detalle de la Compra/Servicio"
         Top             =   1680
         Width           =   4005
      End
      Begin MSDataListLib.DataCombo CboTipoIngreso 
         Height          =   330
         Left            =   4425
         TabIndex        =   11
         ToolTipText     =   "Tipo de Compra/Servicio"
         Top             =   1305
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtSerie 
         Alignment       =   2  'Center
         BackColor       =   &H00F0EADB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3945
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Serie del Documento"
         Top             =   150
         Width           =   855
      End
      Begin VB.TextBox TxtCorrelativo 
         Alignment       =   2  'Center
         BackColor       =   &H00F0EADB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4815
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Correlativo del Documento"
         Top             =   150
         Width           =   1425
      End
      Begin VB.TextBox TxtRUC 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Identificador Tributario"
         Top             =   540
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo CboArea 
         Height          =   330
         Left            =   1050
         TabIndex        =   10
         ToolTipText     =   "Area Destino"
         Top             =   1280
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo CboMoneda 
         Height          =   330
         Left            =   8895
         TabIndex        =   12
         ToolTipText     =   "Moneda de Facturación del Documento"
         Top             =   1305
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboOperacion 
         Height          =   330
         Left            =   6930
         TabIndex        =   14
         ToolTipText     =   "Tipo de Compra/Servicio"
         Top             =   1305
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo CboTipoDocumento 
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         ToolTipText     =   "Tipo de Compra/Servicio"
         Top             =   150
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcLibroCompra 
         Height          =   330
         Left            =   6375
         TabIndex        =   97
         ToolTipText     =   "Tipo de Compra/Servicio"
         Top             =   1665
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpAnio 
         Height          =   330
         Left            =   7410
         TabIndex        =   102
         Top             =   915
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   54788099
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin VB.Label LabelMesContable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Contable:"
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   6210
         TabIndex        =   100
         Top             =   975
         Width           =   1080
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H0080C0FF&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   6150
         Shape           =   4  'Rounded Rectangle
         Top             =   870
         Width           =   3615
      End
      Begin VB.Label lblLibroCompra 
         AutoSize        =   -1  'True
         Caption         =   "Libro de Compra:"
         Height          =   210
         Left            =   5130
         TabIndex        =   98
         Top             =   1695
         Width           =   1230
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   " Transporte:"
         Height          =   210
         Index           =   6
         Left            =   6105
         TabIndex        =   96
         Top             =   945
         Width           =   885
      End
      Begin VB.Label lblFechCaducidad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Caducidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   11670
         TabIndex        =   94
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label lblMontoControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto de Control:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10110
         TabIndex        =   93
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Ingreso:"
         Height          =   210
         Left            =   3225
         TabIndex        =   91
         Top             =   1335
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   210
         Left            =   6120
         TabIndex        =   90
         Top             =   1335
         Width           =   795
      End
      Begin VB.Label LblPDoc 
         Alignment       =   2  'Center
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3660
         TabIndex        =   1
         ToolTipText     =   "Tipo de Documento"
         Top             =   150
         Width           =   285
      End
      Begin VB.Label LblUsuario 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   12180
         TabIndex        =   57
         ToolTipText     =   "Usuario que registra el Documento"
         Top             =   1335
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   210
         Left            =   9810
         TabIndex        =   56
         Top             =   945
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ingreso:"
         Height          =   210
         Left            =   8115
         TabIndex        =   55
         Top             =   570
         Width           =   585
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emisión:"
         Height          =   210
         Left            =   8130
         TabIndex        =   54
         Top             =   195
         Width           =   585
      End
      Begin VB.Label LblEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   10365
         TabIndex        =   53
         ToolTipText     =   "Estado del Documento"
         Top             =   930
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Id. Tributario:"
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   47
         Top             =   570
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cambio:"
         Height          =   210
         Left            =   10215
         TabIndex        =   40
         Top             =   195
         Width           =   570
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   39
         Top             =   945
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   210
         Left            =   525
         TabIndex        =   38
         Top             =   1725
         Width           =   465
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Área:"
         Height          =   210
         Left            =   585
         TabIndex        =   37
         Top             =   1335
         Width           =   405
      End
      Begin VB.Label LblTipoCambio 
         Caption         =   "0.000"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   10830
         TabIndex        =   36
         ToolTipText     =   "Tipo de Cambio de la fecha de emisión"
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N°:"
         Height          =   210
         Left            =   11685
         TabIndex        =   35
         Top             =   195
         Width           =   210
      End
      Begin VB.Label LblCorrelativo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   11985
         TabIndex        =   34
         ToolTipText     =   "Número Correlativo en el mes"
         Top             =   150
         Width           =   1395
      End
      Begin VB.Label LblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   195
         Width           =   855
      End
      Begin VB.Label LblMoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   210
         Index           =   2
         Left            =   8205
         TabIndex        =   32
         Top             =   1335
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   13485
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7470
      Width           =   13545
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   520
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   13605
         TabIndex        =   63
         Top             =   20
         Width           =   13665
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   5
            Left            =   12990
            Picture         =   "frmDocumentoCompra.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   4
            Left            =   12600
            Picture         =   "frmDocumentoCompra.frx":0A50
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   3
            Left            =   12210
            Picture         =   "frmDocumentoCompra.frx":0F92
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   2
            Left            =   780
            Picture         =   "frmDocumentoCompra.frx":14D4
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   1
            Left            =   390
            Picture         =   "frmDocumentoCompra.frx":1A16
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   470
            Index           =   0
            Left            =   0
            Picture         =   "frmDocumentoCompra.frx":1F58
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   400
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1380
            TabIndex        =   64
            Top             =   120
            Width           =   8970
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   1
         Left            =   3360
         MousePointer    =   99  'Custom
         Picture         =   "frmDocumentoCompra.frx":249A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   -540
         Width           =   1080
      End
   End
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5445
      Left            =   30
      TabIndex        =   30
      Top             =   1995
      Width           =   13485
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Impuestos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   10
         Left            =   10200
         Picture         =   "frmDocumentoCompra.frx":259C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4845
         UseMaskColor    =   -1  'True
         Width           =   1080
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Pago rapido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   12210
         MousePointer    =   99  'Custom
         Picture         =   "frmDocumentoCompra.frx":2B26
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2850
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   720
         Top             =   1380
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdObservacion 
         Caption         =   "Observación"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   12210
         Picture         =   "frmDocumentoCompra.frx":30B0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Muestra el motivo de la Anulación del Documento"
         Top             =   2130
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdAdicion 
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
         Height          =   480
         Index           =   3
         Left            =   12210
         Picture         =   "frmDocumentoCompra.frx":363A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancela todo los datos ingresados del Documento"
         Top             =   1665
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdAdicion 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   12210
         Picture         =   "frmDocumentoCompra.frx":3BC4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Elimina un item del detalle"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdAdicion 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   12210
         Picture         =   "frmDocumentoCompra.frx":3CC6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Modifica un item ingresado"
         Top             =   680
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdAdicion 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   12210
         Picture         =   "frmDocumentoCompra.frx":3DC8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Agrega un nuevo item al detalle"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.Frame FraDobserva 
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   3855
         TabIndex        =   68
         Top             =   945
         Visible         =   0   'False
         Width           =   5355
         Begin VB.Frame Frame9 
            Height          =   2520
            Left            =   0
            TabIndex        =   69
            Top             =   -30
            Width           =   5235
            Begin VB.TextBox txtDobserva 
               Height          =   1425
               Left            =   75
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   75
               Text            =   "frmDocumentoCompra.frx":42FA
               Top             =   420
               Width           =   5055
            End
            Begin VB.CommandButton cmdDoaceptar 
               Caption         =   "Salir"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   520
               Left            =   3225
               Picture         =   "frmDocumentoCompra.frx":4302
               Style           =   1  'Graphical
               TabIndex        =   70
               ToolTipText     =   "Sale de esta ventana de Observación"
               Top             =   1920
               UseMaskColor    =   -1  'True
               Width           =   1170
            End
            Begin VB.CommandButton cmdDocancelar 
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
               Height          =   525
               Left            =   2070
               Picture         =   "frmDocumentoCompra.frx":43F4
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Cancela el ingreso del motivo de la Observación"
               Top             =   1920
               UseMaskColor    =   -1  'True
               Width           =   1170
            End
            Begin VB.CommandButton cmdDograbar 
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
               Height          =   520
               Left            =   915
               Picture         =   "frmDocumentoCompra.frx":44F6
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "Graba la Observación de la Anulación"
               Top             =   1920
               UseMaskColor    =   -1  'True
               Width           =   1170
            End
            Begin VB.Label LblAviso 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Index           =   4
               Left            =   0
               TabIndex        =   73
               Top             =   30
               Width           =   5355
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   2550
            Left            =   120
            TabIndex        =   74
            Top             =   150
            Width           =   5235
         End
      End
      Begin TrueOleDBGrid80.TDBGrid GrdGrilla 
         Height          =   2025
         Left            =   75
         TabIndex        =   62
         Top             =   150
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   3572
         _LayoutType     =   4
         _RowHeight      =   19
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   4
         Caption         =   "Detalle del Documento"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         ScrollTrack     =   -1  'True
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
         _StyleDefs(7)   =   ":id=1,.appearance=0,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(8)   =   ":id=1,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=1,.fontname=Arial"
         _StyleDefs(10)  =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(11)  =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=Arial"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.bgpicMode=2,.borderSize=1"
         _StyleDefs(14)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.locked=-1,.bold=-1"
         _StyleDefs(33)  =   ":id=19,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(34)  =   ":id=19,.fontname=Arial"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=34,.fontname=Arial"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(57)  =   ":id=36,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(58)  =   ":id=36,.fontname=Arial"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=825"
         _StyleDefs(63)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=38,.fontname=Arial"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34,.wraptext=-1"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtICE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3915
         TabIndex        =   88
         Top             =   4485
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   5730
         Picture         =   "frmDocumentoCompra.frx":4838
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4845
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1160
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
         Height          =   520
         Index           =   4
         Left            =   11280
         MousePointer    =   99  'Custom
         Picture         =   "frmDocumentoCompra.frx":493A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4845
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.Label LblTotImp3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3735
         TabIndex        =   52
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1095
      End
      Begin VB.Label LblICE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3915
         TabIndex        =   87
         Top             =   4485
         Width           =   945
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Gran Total"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8445
         TabIndex        =   83
         Top             =   4275
         Width           =   1275
      End
      Begin VB.Label lblGranTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8550
         TabIndex        =   82
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   8595
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Label lblLeyD 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblLeyD"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7290
         TabIndex        =   81
         Top             =   4275
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblTotLey1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7320
         TabIndex        =   80
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   7365
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Redondeo"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4860
         TabIndex        =   79
         Top             =   4275
         Width           =   1155
      End
      Begin VB.Label lblRedondeo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4920
         TabIndex        =   78
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   4965
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Label LblItems 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   12915
         TabIndex        =   77
         Top             =   2625
         Width           =   270
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Items:"
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   2
         Left            =   12375
         TabIndex        =   76
         Top             =   2625
         Width           =   435
      End
      Begin VB.Label LblSumDescuento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   11010
         TabIndex        =   65
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Descuento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   11010
         TabIndex        =   67
         Top             =   4275
         Width           =   1155
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   9780
         TabIndex        =   66
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Label LblTotImp1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1395
         TabIndex        =   50
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1095
      End
      Begin VB.Label LblTotImp2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2565
         TabIndex        =   51
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1095
      End
      Begin VB.Label LblImpt3 
         Alignment       =   1  'Right Justify
         Caption         =   "Impuesto 3"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3720
         TabIndex        =   49
         Top             =   4275
         Width           =   1110
      End
      Begin VB.Label LblImpt2 
         Alignment       =   1  'Right Justify
         Caption         =   "Impuesto 2"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2535
         TabIndex        =   48
         Top             =   4275
         Width           =   1110
      End
      Begin VB.Label LblImpt1 
         Alignment       =   1  'Right Justify
         Caption         =   "Impuesto 1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1335
         TabIndex        =   46
         Top             =   4275
         Width           =   1185
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5970
         TabIndex        =   45
         Top             =   4275
         Width           =   1275
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "SubTotal"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   44
         Top             =   4275
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   10545
         TabIndex        =   43
         Top             =   4275
         Width           =   405
      End
      Begin VB.Label LblSubTot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   180
         TabIndex        =   42
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6090
         TabIndex        =   41
         ToolTipText     =   "Haga clic aquí para ajustar el valor..."
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Shape Shape0 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   225
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   1440
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   2610
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   3780
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   6135
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   9825
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   11055
         Top             =   4530
         Width           =   1155
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         Height          =   240
         Left            =   3945
         Top             =   4530
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "I.C.E."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3975
         TabIndex        =   86
         Top             =   4275
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmDocumentoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Cabecera

Public nCorrela As String
Public xPUni As Double
Public sAreaO As String
Public lObserva As Boolean

'DETP 27/02/2008
Public lGrabo           As Boolean
Public lDatosExternos   As Boolean
Public sOCCorrelativo   As String

Dim xCodigo             As String
Dim xArea               As String
Dim xItem               As String
Dim sCosto              As String
Dim xMes, xAnho         As String
Dim xPeriodo            As String
Dim sPrefijo            As String
Public sRefResumida     As String
Public RsDocumentoD     As Recordset
Dim RsBuscaRef          As Recordset
Dim rsTipoCambio        As Recordset   'Cambio
Dim RsCambio            As Recordset
Dim Rs                  As Recordset   'Estado
Dim rsTDocumento        As Recordset   'Combos
Dim RsRDocumento        As Recordset   'Combos, Documento de Referencia
Dim RsProveedor         As Recordset
Dim RsRuc               As Recordset
Dim RsMoneda            As Recordset
Dim RsProgramaPago      As Recordset
Dim RsLibroCompra       As Recordset
Dim RsArea              As Recordset
Dim rsReferencia        As Recordset
Dim RsTingreso          As Recordset
Dim RsTImpuesto         As Recordset
Dim RsSearch            As Recordset
Dim RsMensaje           As Recordset
Dim RsBusca             As Recordset
Dim RsHistorial         As Recordset
Dim RsOperacion         As ADODB.Recordset
Dim nPos                As Integer
Dim iRow                As Integer
Dim nColumna            As Integer
Dim sReqDoc             As String

'Detalle
Dim Tiempo              As Integer
Dim RowC                As Integer
Dim sCodi, sUnid, sCant, sNeto, sPVen, sPuni, sTota, sDesc, sReca As String, sTipoIng As String
Dim sitem               As Double
Dim dRebajaMerma        As Double
Dim dMermaFija          As Double
Dim dMermaArt           As Double
Dim dStockArt           As Double
Dim wDesc               As String
Dim wReca               As String
Public sCerrado         As String
'Public AsumirIUE_IT     As Boolean
'Public swPasar          As Boolean
Public sDtemporal       As String
Public RsDtemporal      As New Recordset
Public lCriterio        As Boolean
Dim swObliga            As Boolean
Dim sResponde           As String
Dim swMover             As Boolean
Dim lOpc                As Boolean
Dim lNavegar            As Boolean
Dim lLoad               As Boolean
Dim RsTemp              As Recordset
Dim iRegs               As Integer
Dim sPregunta           As String

Dim RsUltimoKardex      As Recordset
Dim RsUltimosKardex     As Recordset
Dim RsUltimo            As Recordset
Dim RsUltPromedio       As Recordset
Dim RsUltPromedio2       As Recordset
Dim RsPCosto            As Recordset
Dim RsAnula             As Recordset
Dim RsBRequerimiento    As Recordset
Dim RsSalida            As Recordset
Dim RsUltRequerimiento  As Recordset
Dim RsNotaSalida        As Recordset
Dim RsGEncargado        As Recordset
Dim RsPromedio          As Recordset
'Dim RsUltPromedio          As Recordset
Dim RsSubStock          As Recordset
Dim sDocumento          As String
Dim sTipoProd           As String
Dim nDItem              As String
Dim sArti               As String
Dim lsNotaSalida        As String
Dim sUltimo             As String
Dim sUltPromedio        As String
Dim sPCosto             As String
Dim sPromedio           As String
Dim sSub_Stock          As String
Dim lsSalida            As String
Dim sBRequerimiento     As String
Dim sValor              As String
Dim sAnteriorKD         As String
Dim sAnteriorSKD        As String
Dim sEncargado          As String
Dim sUltRequerimiento   As String
Dim lsAnula             As String
Dim CorrKD              As Variant
Dim CorrSKD             As Variant
Dim CorrSTK             As Variant
Dim RsCorrKD            As Recordset
Dim RsCorrSKD           As Recordset
Dim RsCorrSTK           As Recordset
'Dim Reporte             As New CrtDocumentos
Dim RsReporte           As Recordset
Dim lActualiza          As Boolean
Dim lValidaProducto   As Boolean '- nuevo
Dim qCboTipoDocumento   As String
Dim qLblPDoc            As String
Dim qTxtSerie           As String
Dim qTxtCorrelativo     As String
Dim qDtpEmision         As String
Dim qDtpIngreso         As String
Dim qTxtRUC             As String
Dim qTxtNombre          As String
Dim qTxtProveedor       As String
Dim qChkPrograma        As String
Dim qTxtDiasCredito     As String
Dim qDtpfPrograma       As String
Dim qCboTipoIngreso     As String
Dim qCboMoneda          As String
Dim qCboReferencia      As String
Dim qLblPDocRef         As String
Dim qTxtSerieRef        As String
Dim qTxtCorrelativoRef  As String
Dim qTxtGlosaDocumento  As String
Dim fFechaProcesa       As String
Dim sProValida          As String
Dim nSaldo              As Double, nRedondeo As Double, nLey1 As Double, nLeyAD As Double, nTotal As Double, nCount As Double
Dim mdDocumento         As ClsDocumento
Dim mDocObservado       As ClsDocumento
Dim mmPago              As ClsDocumento
Dim lExporta            As Boolean
Dim clsmDocumento       As ClsDocumento
Dim oDiaContable As clsDiaContable
Dim clsAlmac       As clsAlmacen
Dim nPercepcion         As Double
Dim nTotalDoc           As Double
Dim nGranTotal          As Double
Public vICE             As Double
Dim vICEAntR            As Double
Dim wICE                As Boolean
Public wAfectoICE      As Boolean
Public lCambioDetalle As Boolean
Public wVaRegCompra As Boolean

Dim lImpuesto As Boolean

Dim sCorrecto As String

Dim sFlag As Boolean

Dim RsDatos As Recordset

Dim rsPlantilla As ADODB.Recordset
Dim sCodPorciona As String
'Dim clsPorciona As New clsPorcionamiento
Dim lModificado As Boolean

Dim lCanje As Boolean
'Dim cHistorial As New clsHistorial
Dim lAuditoria As Boolean
'Public lDocCambio As Boolean

Public sxTipoDocumento As String

'Jesus 16/03/16
Public sTempoInsumos As String


'Jesus 16/07/08
Dim lModificarDetalle As Boolean

Dim TipoDocumentoExterno As String

Dim lActualizaImpuesto As Boolean

Dim lModifica As Boolean

Dim meses(11) As String
Dim anios(51)   As String
Dim FechaTransaccion      As String

'-------------------------
Dim lAgregaItem As Boolean
Dim swEntrar As Boolean
Dim nItemEliminar As String
Dim iTBox As Integer
Dim lDesc As String
Dim sUsuarioAutoriza As String
Dim sUsuarioE As String
Dim xDtpEmision As Boolean
Dim xDtpIngreso As Boolean
Dim lMotivoConsumo As Boolean
Dim sBuscar As String

Dim wImp1, wImp2, wImp3, wSubTotal As Double

Dim tPagoRapidoTipoPago As String
Dim sPrf As String
Dim lGeneraNS As Boolean
Sub ActualizaDetalle()
    If RsDtemporal.RecordCount > 0 Then RsDtemporal.MoveFirst
    Dim imp1 As Double
    Do While Not RsDtemporal.EOF
        
        sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "0")
        sUnid = IIf(Not IsNull(RsDtemporal!tUnidadEntrada), RsDtemporal!tUnidadEntrada, "")
        sitem = IIf(Not IsNull(RsDtemporal!nItem), RsDtemporal!nItem, "")
        sNeto = IIf(Not IsNull(RsDtemporal!nNeto), RsDtemporal!nNeto, "")
        sDesc = IIf(Not IsNull(RsDtemporal!nDescuento), RsDtemporal!nDescuento, "")
        sReca = IIf(Not IsNull(RsDtemporal!nRecargo), RsDtemporal!nRecargo, "")
        sPVen = IIf(Not IsNull(RsDtemporal!nPVenta), RsDtemporal!nPVenta, "")
        sPuni = IIf(Not IsNull(RsDtemporal!nPrecio), RsDtemporal!nPrecio, "")
        sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
        sTota = IIf(Not IsNull(RsDtemporal!nTotal), RsDtemporal!nTotal, "")
        sTipoIng = IIf(Not IsNull(RsDtemporal!tTipoIngreso), RsDtemporal!tTipoIngreso, "")
        imp1 = IIf(Not IsNull(RsDtemporal!nPorcentaje1), RsDtemporal!nPorcentaje1, 0)
       'ojoooooooooooooooooooooooooooo lg 24/05/11
       
'        If lModAuditoria Then
'            lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "DDOCUMENTO", "INGRESO DE DOCUMENTOS - DETALLE", "01", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), LblCorrelativo, sItem, "tCorrelativo", "Correlativo", LblCorrelativo, "nItem", "Item", sItem, _
'                                         "nCantidad", "Cantidad", CDbl(sCant), "nPrecio", "Precio", CDbl(sPuni), "nPVenta", "Precio de Venta", CDbl(sPVen), "nDescuento", "Monto de Descuento", CDbl(sDesc), "nRecargo", "Monto de Recargo", CDbl(sReca), "nImpuesto1", "Monto Impuesto 1", CDbl(GrdGrilla.Columns(2).value), "nImpuesto2", "Monto Impuesto 2", CDbl(GrdGrilla.Columns(3).value), "nImpuesto3", "Monto Impuesto 3", CDbl(GrdGrilla.Columns(4).value))
'
'            If lAuditoria = False Then
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'        End If
       
        clsmDocumento.InsdDocumento Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo), TxtProveedor, sitem, _
                                    sCodi, sUnid, CDbl(sCant), CDbl(sPuni), CDbl(sPVen), CDbl(sNeto), _
                                    CDbl(sDesc), CDbl(sReca), CDbl(grdGrilla.Columns(2).value), _
                                    CDbl(grdGrilla.Columns(3).value), _
                                    CDbl(grdGrilla.Columns(4).value), _
                                    IIf(CDbl(grdGrilla.Columns(2).value) > 0, IIf(pais = "005", imp1, CDbl(nPorcentaje1)), 0), _
                                    IIf(CDbl(grdGrilla.Columns(3).value) > 0, CDbl(nPorcentaje2), 0), _
                                    IIf(CDbl(grdGrilla.Columns(4).value) > 0, CDbl(nPorcentaje3), 0), _
                                    CDbl(sTota), "01", IIf(IsNull(CboTipoIngreso.BoundText), "", CboTipoIngreso.BoundText), _
                                    IIf(CDbl(grdGrilla.Columns(2).value) > 0, 1, 0), _
                                    IIf(CDbl(grdGrilla.Columns(3).value) > 0, 1, 0), _
                                    IIf(CDbl(grdGrilla.Columns(4).value) > 0, 1, 0), LblCorrelativo, _
                                    IIf(IsNull(RsDtemporal!TCUENTACONTABLE), "", RsDtemporal!TCUENTACONTABLE), _
                                    IIf(IsNull(RsDtemporal!tCentroCosto), "", RsDtemporal!tCentroCosto), IIf(IsNull(RsDtemporal!tAreaOrigen), "", RsDtemporal!tAreaOrigen), "I", IIf(IsNull(RsDtemporal!lAsumirIUE_IT), 0, RsDtemporal!lAsumirIUE_IT), _
                                    IIf(IsNull(RsDtemporal!nRetencion1), 0, RsDtemporal!nRetencion1), _
                                    IIf(IsNull(RsDtemporal!nRetencion2), 0, RsDtemporal!nRetencion2), _
                                    IIf(IsNull(RsDtemporal!nRetencion3), 0, RsDtemporal!nRetencion3), _
                                    IIf(IsNull(RsDtemporal!nPRetencion1), 0, RsDtemporal!nPRetencion1), _
                                    IIf(IsNull(RsDtemporal!nPRetencion2), 0, RsDtemporal!nPRetencion2), _
                                    IIf(IsNull(RsDtemporal!nPRetencion3), 0, RsDtemporal!nPRetencion3), _
                                    IIf(IsNull(RsDtemporal!lIncluidoRetencion1), 0, RsDtemporal!lIncluidoRetencion1), _
                                    IIf(IsNull(RsDtemporal!lIncluidoRetencion2), 0, RsDtemporal!lIncluidoRetencion2), _
                                    IIf(IsNull(RsDtemporal!lIncluidoRetencion3), 0, RsDtemporal!lIncluidoRetencion3), _
                                    IIf(IsNull(RsDtemporal!nICE), 0, RsDtemporal!nICE), IIf(IsNull(RsDtemporal!nOtrosCargosInafecto), 0, RsDtemporal!nOtrosCargosInafecto)

        'Guarda la percepcion en el detalle del documento - ddocumento
        'cambio jesus 2014-09-11
        If nLey1 > 0 Then
            Isql = "Exec spSavePerceptionDetails '" & LblCorrelativo & "'"
            CnAlmacen.Execute Isql
        End If
        'fin cambio
        
        'JESUS
        If nLeyAD > 0 Then
            Isql = "Exec spSaveLeyADDetails '" & LblCorrelativo & "'"
            CnAlmacen.Execute Isql
        End If
        '------------02.11.2017-----------------

' asi era ::::::   ' lg RsDtemporal!tCuentacontable, RsDtemporal!tCentroCosto,"I"

       RsDtemporal.MoveNext
    Loop
End Sub

Sub BotonesDetalle(lSw As Boolean)
    CmdAdicion(0).Enabled = lSw
    'CmdAdicion(1).Enabled = lSw
    CmdAdicion(2).Enabled = lSw
End Sub

Sub Liberar()
    Set rsTDocumento = Nothing
    Set RsRDocumento = Nothing
    Set RsBusca = Nothing
    Set RsDocumentoD = Nothing
    Set RsMoneda = Nothing
    Set RsArea = Nothing
    Set rsReferencia = Nothing
    Set RsTingreso = Nothing
    Set RsUltimoKardex = Nothing
    Set RsTImpuesto = Nothing
    Set RsUltimosKardex = Nothing
    Set Rs = Nothing
    Set RsUltimo = Nothing
    Set RsSearch = Nothing
    Set RsProveedor = Nothing
    Set RsUltPromedio = Nothing
    Set RsRuc = Nothing
    Set RsBRequerimiento = Nothing
    Set RsPCosto = Nothing
    Set RsCorrKD = Nothing
    Set RsPromedio = Nothing
    Set RsAnula = Nothing
    Set RsCorrSKD = Nothing
    Set RsMensaje = Nothing
    Set RsSalida = Nothing
    Set RsCorrSTK = Nothing
    Set RsUltRequerimiento = Nothing
    Set RsCambio = Nothing
    
    'If Len(Trim(sDtemporal)) <> "" Then
    If Len(Trim(sDtemporal)) > 0 Then
        Isql = "Drop Table " & sDtemporal & ""
        CnAlmacen.Execute Isql
    End If
End Sub

Sub LlenaGrilla()
    Isql = "sp_TraeDetalleDocumento '" & LblCorrelativo & "'"
    Set RsDocumentoD = Lib.OpenRecordset(Isql, CnAlmacen)
    Set grdGrilla.DataSource = RsDocumentoD
    LblItems = Format(RsDocumentoD.RecordCount, "##0")
End Sub

Private Sub LlenaCombos()
            
'    Isql = "Select tCodigoTipoDocumento as Codigo, tDetallado AS Descripcion FROM TTipoDocumento Where lActivo=1 And lInterno='0' ORDER BY Descripcion"
'    Llena_Combo Isql, CboReferencia, RsRDocumento, "Codigo", "Descripcion"
'
'    Isql = "Select tresumido as Codigo, tDetallado AS Descripcion FROM TTipoDocumento Where lActivo=1 And lInterno='0' ORDER BY Descripcion"
'    Llena_Combo Isql, CboReferencia, RsRDocumento, "Codigo", "Descripcion"
    
'    If ErpOF_Activo Then
'        If Sw Then
'            Isql = "Select tCodigoTipoDocumento As Codigo,tDetallado AS Descripcion FROM TTipoDocumento Where lActivo=1 And lInterno='0' And tCodigoTipoDocumento = '" & Xml_tTipoDocPref & "' ORDER BY Descripcion"
'        Else
'            Isql = "Select tCodigoTipoDocumento As Codigo,tDetallado AS Descripcion FROM TTipoDocumento Where lActivo=1 And lInterno='0' ORDER BY Descripcion"
'        End If
'    Else
        Isql = "Select tCodigoTipoDocumento As Codigo,tDetallado AS Descripcion FROM TTipoDocumento Where lActivo=1 And lInterno='0' ORDER BY Descripcion"
    'End If
    Llena_Combo Isql, CboTipoDocumento, rsTDocumento, "Codigo", "Descripcion"
    
    Isql = "SELECT Codigo,Descripcion FROM vMoneda ORDER BY Descripcion"
    Llena_Combo Isql, CboMoneda, RsMoneda, "Codigo", "Descripcion"
    
'    Isql = "Select tCodigo As Codigo,tdetallado As Descripcion From tTabla Where tTabla='PROGRAMAPAGO' And lActivo='1'"
'    Llena_Combo Isql, dtcProgramaPago, RsProgramaPago, "Codigo", "Descripcion"

    
'    If lExtensionAlmacen Then
'        If CBool(Calcular("Select (case when dbo.TUSUARIO.tGrupoUsuario='00' then 1 else Isnull(dbo.tPermisoGrupo.lOpcion11,0) END) As Codigo FROM dbo.TUSUARIO LEFT OUTER JOIN dbo.tPermisoGrupo ON dbo.TUSUARIO.tGrupoUsuario = dbo.tPermisoGrupo.tGrupoUsuario Where dbo.TUSUARIO.lactivo=1 and (dbo.TUSUARIO.TUSUARIO = '" & sUsuario & "')", CnAlmacen)) = False Then
'            If lMotivoConsumo Then
'                Isql = "SELECT dbo.vArea.Codigo, dbo.vArea.Descripcion " & _
'                       "FROM dbo.vArea INNER JOIN dbo.TSEDEALMACEN ON dbo.vArea.Codigo = dbo.TSEDEALMACEN.tCodigoArea " & _
'                       "WHERE (dbo.vArea.lActivo = 1) AND ((dbo.TSEDEALMACEN.tCodigoSedeAlmacen = N'" & tCodigoSede & "' AND (dbo.TSEDEALMACEN.lSedeAlmacen = 0)) or dbo.vArea.Codigo='" & tCodigoSede & "')  And dbo.vArea.Codigo='000' " & _
'                       "ORDER BY dbo.vArea.Codigo"
'            Else
'                Isql = "SELECT dbo.vArea.Codigo, dbo.vArea.Descripcion " & _
'                       "FROM dbo.vArea INNER JOIN dbo.TSEDEALMACEN ON dbo.vArea.Codigo = dbo.TSEDEALMACEN.tCodigoArea " & _
'                       "WHERE (dbo.vArea.lActivo = 1) AND ((dbo.TSEDEALMACEN.tCodigoSedeAlmacen = N'" & tCodigoSede & "' AND (dbo.TSEDEALMACEN.lSedeAlmacen = 0)) or dbo.vArea.Codigo='" & tCodigoSede & "') " & _
'                       "ORDER BY dbo.vArea.Codigo"
'            End If
'        Else
'            If lMotivoConsumo Then
'                Isql = "Select tCodigo As Codigo,tResumido As Descripcion From tTabla Where tTabla='Area' And lActivo='1' And tCodigo='000'"
'            Else
'                Isql = "Select tCodigo As Codigo,tResumido As Descripcion From tTabla Where tTabla='Area' And lActivo='1'"
'            End If
'        End If
'    Else
'        If lMotivoConsumo Then
'            Isql = "Select tCodigo As Codigo,tResumido As Descripcion From tTabla Where tTabla='Area' And lActivo='1' And tCodigo='000'"
'        Else
'            Isql = "Select tCodigo As Codigo,tResumido As Descripcion From tTabla Where tTabla='Area' And lActivo='1'"
'        End If
'    End If
    
    Isql = "Select tCodigo As Codigo,tResumido As Descripcion From tTabla Where tTabla='Area' And lActivo='1' "
    
    Llena_Combo Isql, CboArea, RsArea, "Codigo", "Descripcion"
    
    Isql = " SELECT Codigo As Codigo,Descripcion As Descripcion FROM vTipoIngreso WHERE LACTIVO=1 ORDER BY Descripcion"
    Llena_Combo Isql, CboTipoIngreso, RsTingreso, "Codigo", "Descripcion"

'    If ErpSG_Activo Then
'        Isql = "SELECT TCODIGO As Codigo, TRESUMIDO As Descripcion FROM tTabla WHERE tTabla = 'tiposervicio' AND lActivo = '1'"
'        Llena_Combo Isql, cboOperacion, RsOperacion, "Codigo", "Descripcion"
'    Else
'        Isql = "SELECT TCODIGO As Codigo, TRESUMIDO As Descripcion FROM tTabla WHERE tTabla = 'tOperacion' AND lActivo = '1'"
'        Llena_Combo Isql, cboOperacion, RsOperacion, "Codigo", "Descripcion"
'    End If
        
    Isql = "SELECT TCODIGO As Codigo, TRESUMIDO As Descripcion FROM tTabla WHERE tTabla = 'tOperacion' AND lActivo = '1'"
    Llena_Combo Isql, cboOperacion, RsOperacion, "Codigo", "Descripcion"

    Isql = "Select tCodigo As Codigo,tdetallado As Descripcion From tTabla Where tTabla='LIBRO_COMPRA' And lActivo='1'"
    Llena_Combo Isql, dtcLibroCompra, RsLibroCompra, "Codigo", "Descripcion"

End Sub

Private Sub LlenaMesContable()
    Dim i As Integer
    Dim j As Integer
    
        CmbMes.Clear
        For i = 0 To 11
            CmbMes.AddItem meses(i)
        Next i
        CmbMes.ListIndex = Month(DtpIngreso.value) - 1

End Sub

'Sub Asignar()
'    With frmDocumento.RsCabecera
'        If .EOF Then Exit Sub
'
'        sReqDoc = IIf(IsNull(!tRequerimiento), "", Trim(!tRequerimiento))
'        swMover = True
'        If pais <> "002" Then
'            TxtSerie = Mid(Trim(!tDocumento), 2, 5)
'            TxtCorrelativo = Mid(Trim(!tDocumento), 7, 9)
'        Else
'            TxtSerie = Mid(Trim(!tDocumento), 2, 6)
'            TxtCorrelativo = Mid(Trim(!tDocumento), 8, 8)
'        End If
'        TxtGlosaDocumento = UCase(IIf(IsNull(!CabGlosa), "", Trim(!CabGlosa)))
'
'        If pais <> "002" Then
'            TxtSerieRef = Mid(Trim(!tDocumentoReferencia), 2, 5)
'            TxtCorrelativoRef = Mid(Trim(!tDocumentoReferencia), 7, 9)
'        Else
'            TxtSerieRef = Mid(Trim(!tDocumentoReferencia), 2, 6)
'            TxtCorrelativoRef = Mid(Trim(!tDocumentoReferencia), 8, 8)
'        End If
'        TxtProveedor = IIf(IsNull(!tcodigoproveedor), "", Trim(!tcodigoproveedor))
'        TxtNombre = IIf(IsNull(!RazonSocial), "", Trim(!RazonSocial))
'        TxtRUC = IIf(IsNull(!tIndicadorTributario), "", Trim(!tIndicadorTributario))
'        TxtFPago = IIf(IsNull(!FecCancelado), "", Format(!FecCancelado, "dd-mm-yyyy HH:mm"))
'        If !fProcesa = "01/01/1900" Then
'            txtFechaProceso = ""
'        Else
'            txtFechaProceso = IIf(IsNull(!fProcesa), "", Format(!fProcesa, "dd/mm/yyyy HH:mm"))
'        End If
'
'        CboReferencia.BoundText = IIf(IsNull(!tTipoDocumentoRef), "", Trim(!tTipoDocumentoRef))
'        CboArea.BoundText = IIf(IsNull(!tcodigoarea), "", Trim(!tcodigoarea))
'        CboMoneda.BoundText = IIf(IsNull(!tMoneda), "", Trim(!tMoneda))
'        CboTipoIngreso.BoundText = UCase(IIf(IsNull(!CabTipoIngreso), "", Trim(!CabTipoIngreso)))
'
'        dtcProgramaPago.BoundText = UCase(IIf(IsNull(!tCodProgramaPago), "", Trim(!tCodProgramaPago)))
'
'        DtpEmision = IIf(IsNull(!FecEmision), "", Trim(!FecEmision))
'        DtpIngreso = IIf(IsNull(!FecIngreso), "", Trim(!FecIngreso))
'        DtpFPrograma = IIf(IsNull(!fecprogramacion), DtpIngreso, Trim(!fecprogramacion))
'        TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
'        ChkPrograma = IIf(IsNull(!fecprogramacion), 0, 1)
'        If ChkPrograma Then
''            If lProgramaPagoPre Then
''                dtcProgramaPago.Enabled = True
''            Else
'                DtpFPrograma.Enabled = True
'                TxtDiasCredito.Enabled = True
''            End If
'        Else
''            If lProgramaPagoPre Then
''                dtcProgramaPago.Enabled = False
''            Else
'                DtpFPrograma.Enabled = False
'                TxtDiasCredito.Enabled = False
''            End If
'        End If
'        'JESUS
'        '========
'        lCanje = IIf(IsNull(!lCanje), 0, !lCanje)
'        '========
'        chkDetraccion.value = IIf(IIf(IsNull(!lDetraccion), False, !lDetraccion), 1, 0)
'        LblEstado = UCase(IIf(IsNull(!CabEstado), "", Trim(!CabEstado)))
'        LblCorrelativo = IIf(IsNull(!tCorrelativo), "", Trim(!tCorrelativo))
'        LblPDoc = Left(Trim(!tDocumento), 1)
'        LblPDocRef = Left(Trim(!tDocumentoReferencia), 1)
'        LblEstado = UCase(IIf(IsNull(!CabEstadoDocumento), "", Trim(!CabEstadoDocumento)))
'        LblTipoCambio = Format(IIf(IsNull(!CabCambio), "", !CabCambio), "##0.000")
'        dtcLibroCompra.BoundText = IIf(IsNull(!tLibroCompra), "", !tLibroCompra)
'        dtpAnio.value = IIf(IsNull(!tAnioContable), Year(!FecIngreso), "01/" & Month(!FecIngreso) & "/" & !tAnioContable)
'        CmbMes.ListIndex = IIf(IsNull(!tMesContable), Month(!FecIngreso), !tMesContable) - 1
'        cboOperacion.BoundText = IIf(IsNull(!tCodigoOperacion), "", !tCodigoOperacion)
'        If Left(!tDocumento, 1) = "F" Or Left(!tDocumento, 1) = "T" Then
'            cboOperacion.Enabled = True
'        End If
'
'        If .RecordCount > 0 Then
'            If Not .EOF Then
'                If IsNull(!lDescuento) Then
'                    lDesc = ""
'                Else
'                    If !lDescuento Then
'                        lDesc = "G"
'                    Else
'                        lDesc = "I"
'                    End If
'                End If
'            Else
'                lDesc = ""
'            End If
'        End If
'
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        LblSumDescuento = IIf(IsNull(!CabDescuento), 0, Format(!CabDescuento, "###,##0.00"))
'        LblSubTot = IIf(IsNull(!CabNeto), "", Format(!CabNeto, "###,###,##0.00"))
'        LblTotImp1 = Format(IIf(IsNull(!CabImp1), 0, !CabImp1), "###,###,##0.00")
'        If CDbl(LblTotImp1) = 0 And LblImpt1 = "" Then
'            Shape1.Visible = False
'            LblTotImp1.Visible = False
'        Else
'            Shape1.Visible = True
'            LblTotImp1.Visible = True
'        End If
'        LblTotImp2 = Format(IIf(IsNull(!CabImp2), 0, !CabImp2), "###,###,##0.00")
'        If CDbl(LblTotImp2) = 0 And LblImpt2 = "" Then
'            Shape2.Visible = False
'            LblTotImp2.Visible = False
'        Else
'            Shape2.Visible = True
'            LblTotImp2.Visible = True
'        End If
'        LblTotImp3 = Format(IIf(IsNull(!CabImp3), 0, !CabImp3), "###,###,##0.00")
'        If CDbl(LblTotImp3) = 0 And LblImpt3 = "" Then
'            Shape3.Visible = False
'            LblTotImp3.Visible = False
'        Else
'            Shape3.Visible = True
'            LblTotImp3.Visible = True
'        End If
'        nLey1 = IIf(IsNull(!nPercepcion), 0, !nPercepcion)
'        lblTotLey1 = Format(IIf(IsNull(!nPercepcion), 0, !nPercepcion), "###,###,##0.00")
'        nRedondeo = IIf(IsNull(!nRedondeo), 0, !nRedondeo)
'        lblRedondeo = Format(IIf(IsNull(!nRedondeo), 0, !nRedondeo), "###,###,##0.00")
'        LblTotal = Format(IIf(IsNull(!TOTAL), 0, !TOTAL), "###,###,##0.00")
'        txtTransporte.Text = Format(IIf(IsNull(!nTransporte), 0, !nTransporte), "##,###,##0.00")
'        If lMontoControl Then
'            txtMontoControl = Format(IIf(IsNull(!TOTAL), 0, !TOTAL), "###,###,##0.00")
'            txtMontoControl.backColor = &HC0FFC0
'        End If
'        LblSaldo = Format(IIf(IsNull(!CabSaldo), 0, !CabSaldo), "###,###,##0.00")
'        Select Case pais
'            Case "001"
'                LblICE = Format(IIf(IsNull(!nICE), 0, !nICE), "###,###,##0.00")
'                vICE = LblICE.Caption
'                vICEAntR = LblICE.Caption
'                wAfectoICE = !lAfectoICE
'        End Select
'        lCambioDetalle = False
'        nTotalDoc = IIf(IsNull(!TOTAL), 0, !TOTAL)
'        nGranTotal = nTotalDoc + nLey1
'        lblGranTotal = Format(nGranTotal, "###,###,##0.00")
'        If !CabEstadoDocumento = "03" Or !CabEstadoDocumento = "04" Then
'            LblTotImp1.ToolTipText = ""
'            LblTotImp2.ToolTipText = ""
'            LblTotImp3.ToolTipText = ""
'            lblRedondeo.ToolTipText = ""
'            LblTotal.ToolTipText = ""
'        Else
'            LblTotImp1.ToolTipText = "Haga clic aquí para ajustar el valor..."
'            LblTotImp2.ToolTipText = "Haga clic aquí para ajustar el valor..."
'            LblTotImp3.ToolTipText = "Haga clic aquí para ajustar el valor..."
'            lblRedondeo.ToolTipText = "Haga clic aquí para ajustar el valor..."
'            LblTotal.ToolTipText = "Haga clic aquí para ajustar el valor..."
'        End If
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        'Tema para Bolivia
'        txtCodigoControl.Text = IIf(IsNull(!tCodigoControl), "", !tCodigoControl)
'        txtAutorizacion.Text = IIf(IsNull(!tCodigoAutorizacion), "", !tCodigoAutorizacion)
'        dtpFechCaducidad.value = IIf(IsNull(!fFechaCaducidad), Date, !fFechaCaducidad)
'
'        LblUsuario = IIf(IsNull(!tUsuario), "", Trim(!tUsuario))
'        Set Rs = Lib.OpenRecordset("Select Descripcion From vEstadoDocumento Where Codigo ='" & !CabEstadoDocumento & "'", CnAlmacen)
'        Select Case !CabEstadoDocumento
'            Case Is = "01", "02", "03", "04", "05": LblEstado = Rs!descripcion
'        End Select
'        swMover = False
'    End With
'
'    ActivarBotones True
'    ControlEstado
'    txtDobserva = IIf(Not frmDocumento.RsCabecera.EOF, IIf(Not IsNull(frmDocumento.RsCabecera!CabObservacion), frmDocumento.RsCabecera!CabObservacion, ""), "")
'
'    If TxtFPago <> "" Then LblEstado = LblEstado & " (" & TxtFPago & ")"
'End Sub

Private Sub CboArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CboTipoIngreso.SetFocus
End Sub

Private Sub CboMoneda_Change()
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xCboMoneda = True
'    End If
End Sub

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtGlosaDocumento.SetFocus
End Sub

Private Sub CboReferencia_Change()
'    If LblPDocRef <> "" Then
''        If frmDocumento.RsCabecera.RecordCount > 0 Then
''            If frmDocumento.RsCabecera!CabEstadoDocumento = "02" Then    'Procesado
''                TxtSerieRef.Enabled = False
''                LblPDocRef.ForeColor = &H808080
''            Else
''                TxtSerieRef.Enabled = True
''                LblPDocRef.ForeColor = &H800000
''            End If
''        End If
'
'        TxtSerieRef.Enabled = True
'        LblPDocRef.ForeColor = &H800000
'
'    End If
    If cmdOpcion(3).Visible Then
        cmdOpcion(4).Enabled = True
        'xCboReferencia = True
    End If
End Sub



Private Sub CboTipoDocumento_Change()
'    If LblPDoc <> "" Then
'        TxtSerie.Enabled = True
'        CboReferencia.Enabled = True
'    End If
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xCboTipoDocumento = True
'    End If
'    GeneraMensaje
End Sub

Private Sub CboTipoDocumento_Click(Area As Integer)
    Isql = "Select tResumido, lCtaCte, lRegistroCompras  From tTipoDocumento Where tCodigoTipoDocumento='" & CboTipoDocumento.BoundText & "'"
    Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
    If Rs.RecordCount <= 0 Then Exit Sub
    
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If CboTipoDocumento.BoundText = "02" Then       'Recibo por Honorarios
        CboTipoIngreso = "Servicio"
    End If
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    CboTipoDocumento_LostFocus
    LblPDoc = Trim(Rs!tResumido)
End Sub

Private Sub CboTipoDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtSerie.SetFocus
End Sub

Private Sub CboTipoIngreso_Change()
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xCboTipoIngreso = True
'    End If
End Sub

Private Sub cboTipoIngreso_Click(Area As Integer)
    CboMoneda.Enabled = True
    TxtGlosaDocumento.Enabled = True
    txtTransporte.Enabled = True
    DtpEmision.Enabled = True
    
End Sub

Private Sub CboTipoIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CboMoneda.SetFocus
End Sub



Private Sub CboTipoDocumento_LostFocus()
    Dim rsReferencia As Recordset
    Dim sReferencia As String
    'Preguntar si debe estar
    If Sw Then
        If CboTipoDocumento.BoundText = "01" Or CboTipoDocumento.BoundText = "12" Then
            cboOperacion.Enabled = True
        Else
            cboOperacion.BoundText = Calcular("select tCodigo as Codigo from tTabla where tValor='G' and tTabla='TOPERACION' and lActivo=1", CnAlmacen)
            cboOperacion.Enabled = False
        End If
    End If
    Isql = "Select tCodigoTipoDocumento,lReferencia From tTipoDocumento Where tCodigoTipoDocumento='" & Me.CboTipoDocumento.BoundText & "'"
    Set rsReferencia = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not rsReferencia.EOF Then
        sReferencia = IIf(IsNull(rsReferencia!lReferencia), "", rsReferencia!lReferencia)
'        If sReferencia Then
'            If Sw Then
'                CboReferencia.Enabled = True
'                TxtSerieRef.Enabled = True
'                TxtCorrelativoRef.Enabled = True
'            Else
'                CboReferencia.Enabled = False
'                TxtSerieRef.Enabled = False
'                TxtCorrelativoRef.Enabled = False
'            End If
'        Else
'            CboReferencia.Enabled = False
'            TxtSerieRef.Enabled = False
'            TxtCorrelativoRef.Enabled = False
'        End If
'JESUS
'===============================
'        Dim nCantReferencia As Integer
'
'        nCantReferencia = Calcular("select isnull(count(*),0) as Codigo from tReferencia where tReferencia='" & CboTipoDocumento.BoundText & "'", CnAlmacen)
'        If nCantReferencia <> 0 Then
'            Isql = "SELECT TREFERENCIA.tTipoDocumento as Codigo, TTIPODOCUMENTO.tDetallado as Descripcion FROM  TTIPODOCUMENTO INNER JOIN TREFERENCIA ON TTIPODOCUMENTO.tCodigoTipoDocumento = TREFERENCIA.tTipoDocumento WHERE TREFERENCIA.tReferencia='46' and TREFERENCIA.lReferencia='1'"
'            Llena_Combo Isql, CboReferencia, rsReferencia, "Codigo", "Descripcion"
'            If Sw Then       'Nuevo
'                CboReferencia.BoundText = ""
'            End If
'        End If
        
'===============================

'        Isql = "Select tReferencia As Codigo,tNombreReferencia As Descripcion From tReferencia Where tTipoDocumento='" & CboTipoDocumento.BoundText & "' And lReferencia='1'"
'        Llena_Combo Isql, CboReferencia, RsReferencia, "Codigo", "Descripcion"
'        If Sw Then       'Nuevo
'            CboReferencia.BoundText = ""
'        Else             'Modifica
'            CboReferencia.BoundText = IIf(IsNull(FrmDocumento.RsCabecera!tTipoDocumentoRef), "", FrmDocumento.RsCabecera!tTipoDocumentoRef)
'        End If
    End If
    Isql = "Select tResumido, lCtaCte, lRegistroCompras  From tTipoDocumento Where tCodigoTipoDocumento='" & CboTipoDocumento.BoundText & "'"
    Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
    If Rs.RecordCount <= 0 Then Exit Sub
    
   
    
    LblPDoc = Trim(Rs!tResumido)
End Sub

Sub ControlEstado()
'    If LblEstado <> "" And frmDocumento.RsCabecera.RecordCount > 0 Then
'        With frmDocumento.RsCabecera
'            If Not .EOF Then
'                If !CabEstadoDocumento = "01" Then  'Generado
'                    ActivaAdicion True
'                    CmdDescuento.Enabled = True     'Descuento
'                    cmdOpcion(3).Visible = False     'Modificar
'                    cmdOpcion(2).Caption = "Anular"
'                    cmdOpcion(2).Visible = True    'Anular
'                    cmdOpcion(2).Enabled = True    'False     Anular
'                    cmdOpcion(4).Enabled = True     'Grabar
'                    cmdOpcion(8).Enabled = False
'                    cmdPago.Enabled = False         'Pagar
'                    'Bolivia
'                    Select Case pais
'                        Case "000", "005": cmdOpcion(10).Enabled = True
'                        Case "001": cmdOpcion(10).Enabled = False
'                        Case "002": cmdOpcion(10).Enabled = True
'                    End Select
'                    cmdFind.Enabled = False
'                    dtcProgramaPago.Enabled = True
''                    DtpFPrograma.Enabled = True
'                    ChkPrograma.Enabled = True
'                    TxtDiasCredito.backColor = &HFFFFFF          'Habilita
''                    TxtDiasCredito.Enabled = True
'                    If GrdGrilla.VisibleRows > 0 Then
'                        CmdProcesar.Enabled = True
'                    Else
'                        CmdProcesar.Enabled = False
'                    End If
'                    EstadoCabecera True
'                    CmdProcesar.Caption = "Procesar"
'                    cmdOpcion(3).Visible = False
'                    CmdObservacion.Enabled = False
'
'                End If
'
'                If !CabEstadoDocumento = "02" Then  'Procesado
'                    cmdOpcion(2).Enabled = False     'Anular***
'                    cmdOpcion(4).Enabled = False    'Grabar
'                    cmdOpcion(7).Enabled = True     'Canje
'                    If LblPDoc = "G" Or Round(CDbl(LblTotal) + CDbl(lblTotLey1), 3) = Round(CDbl(LblSaldo), 3) Then
'                        cmdOpcion(8).Enabled = False    'Historial
'                        cmdPago.Enabled = False         'Pagar
'                    Else
'                        cmdOpcion(8).Enabled = True     'Historial
'                    End If
'                    cmdPago.Enabled = True          'Pagar
'                    CmdDescuento.Enabled = False    'Descuento
'                    ChkPrograma.Enabled = True
'                    dtcProgramaPago.Enabled = False
'                    TxtDiasCredito.backColor = &HFFFFFF          'Habilita
'                    'TxtDiasCredito.Enabled = True
'                    'DtpFPrograma.Enabled = True
'                    TxtDiasCredito.Enabled = False
'                    DtpFPrograma.Enabled = False
'                    ChkPrograma.Enabled = False
'                    dtcProgramaPago.Enabled = False
'
'                    ActivaAdicion False
'                    EstadoCabecera False
'
'                    CmdProcesar.Enabled = True
'                    CmdProcesar.Caption = "Desprocesar"
'
'                    cmdOpcion(3).Top = cmdOpcion(2).Top
'                    cmdOpcion(3).Left = cmdOpcion(2).Left
'                    cmdOpcion(3).Visible = True                'Botón Modificar
'                    CmdObservacion.Enabled = False
'                End If
'
'                If !CabEstadoDocumento = "03" Then  'Cancelado
'                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                    'Si el documento está cancelado, ya no habilita la fecha de programación...
'                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                    ChkPrograma.Enabled = False
'                    dtcProgramaPago.Enabled = False
'                    TxtDiasCredito.backColor = &H8000000F       'Deshabilita
'                    TxtDiasCredito.Enabled = False
'                    dtcProgramaPago.Enabled = False
'                    DtpFPrograma.Enabled = False
'                    cmdOpcion(2).Enabled = False     'Anular***
'                    cmdOpcion(4).Enabled = False
'                    cmdOpcion(8).Enabled = True
'                    cmdPago.Enabled = False          'Pagar
'                    cmdOpcion(10).Enabled = False
'                    CmdDescuento.Enabled = False
'                    CmdProcesar.Enabled = False
'                    cmdOpcion(3).Top = cmdOpcion(2).Top
'                    cmdOpcion(3).Left = cmdOpcion(2).Left
'                    cmdOpcion(3).Visible = True                'Botón Modificar
'                    CmdObservacion.Enabled = False
'
'                    ActivaAdicion False
'                    EstadoCabecera False
'                   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                End If
'
'                If !CabEstadoDocumento = "04" Then  'Anulado
'                    CmdProcesar.Enabled = False
'                    cmdOpcion(3).Visible = False    'Modificar
'                    cmdOpcion(2).Caption = "Desanular"
'                    cmdOpcion(2).Visible = True    'Anular
'                    cmdOpcion(2).Enabled = True    'False     Anular
'                    cmdOpcion(4).Enabled = False    'Grabar
'                    cmdOpcion(7).Enabled = False    'Canje
'                    cmdOpcion(8).Enabled = False    'Saldos
'                    cmdPago.Enabled = False         'Pagar
'                    CmdDescuento.Enabled = False    'True   'Descuento
'                    ChkPrograma.Enabled = False
'                    dtcProgramaPago.Enabled = False
'                    TxtDiasCredito.backColor = &H8000000F       'Deshabilita
'                    TxtDiasCredito.Enabled = False
'                    DtpFPrograma.Enabled = False
'                    dtcProgramaPago.Enabled = False
'
'                    ActivaAdicion False
'                    EstadoCabecera False
'                    CmdProcesar.Enabled = False
'                    CmdObservacion.Enabled = True
'                End If
'            End If
'        End With
'    End If
    If CboMoneda = "01" Then DtpEmision.Enabled = True
    'DtpIngreso.Enabled = False
End Sub

Private Sub CmdAdicion_Click(Index As Integer)
    
    sxTipoDocumento = CboTipoDocumento.BoundText
    
    'Jesus 08-02-2016
    'Se agrego validacion de estado
    'Jesus 08-07-2016  ** validacion por base de datos
    lModificarDetalle = Calcular("select IsNull(lModificarDetalleDoc,0) as Codigo from tParametro ", CnAlmacen)
    If lModificarDetalle = False Then
    'Fin 08-07-2016  ** validacion por base de datos
        If LblCorrelativo.Caption <> "" Then
            If ValidaDuplicidad Then
                MsgBox "El Documento ha cambiado de estado, debe verificar otras estaciones", vbExclamation, sMensaje
                Exit Sub
            End If
        End If
    End If
    'Fin 08-02-2016
    
    Select Case Index
        
        Case Is = 0  'Adicionar Registro
        
            lNavegar = False
            lLoad = False
            ValidaCierre
            If sCerrado = "Verdadero" Then Exit Sub

'            If lDesc = "G" Then
'                MsgBox "No es posible adicionar más items. Descuento global aplicado...", vbInformation, sMensaje: Exit Sub
'            End If
            If Len(Trim(CboTipoDocumento.BoundText)) = 0 Then
                MsgBox "Debe seleccionar el Tipo de Documento...", vbInformation, sMensaje
                CboTipoDocumento.SetFocus
                Exit Sub
            End If
            If Len(Trim(CboTipoIngreso.BoundText)) = 0 Then
                MsgBox "Debe seleccionar el Tipo de Ingreso...", vbInformation, sMensaje
                CboTipoIngreso.SetFocus
                Exit Sub
            End If
            swEntrar = True
            FechaEmision
            lAgregaItem = True
            
            If swEntrar = True Then
                FrmSeleccionProductos.Adicionar(2).Enabled = False
                FrmSeleccionProductos.lAgregaItem = True
                FrmSeleccionProductos.Caption = "Agregar Artículos"
                FrmSeleccionProductos.Show vbModal
                
                CalculaTmp
            End If
        Case Is = 1 'Modificar Registro
        
             lNavegar = False
            lLoad = False
            ValidaCierre
            If sCerrado = "Verdadero" Then Exit Sub
            
            If grdGrilla.VisibleRows = 0 Then
                MsgBox "No existe datos para Modificar...", vbInformation, sMensaje
                Exit Sub
            End If
            Dim SwTmp As Boolean
            SwTmp = True
            If SwTmp Then           ' Modifica en el Temporal
                If RsDtemporal.RecordCount > 0 Then
                    If grdGrilla.VisibleRows > 0 Then
                        With FrmSeleccionProductos
                            .AsignaTmp
                            .TxtCodigo.Enabled = False
                            '.TxtCodigoBarra.Enabled = False
                            .txtArticulo.Enabled = False
                            '.CmdCreaArticulo.Enabled = False
                            .ChkImp1.Caption = sImpuesto1
                            .ChkImp2.Caption = sImpuesto2
                            .ChkImp3.Caption = sImpuesto3
                            .TxtImp1 = nPorcentaje1
                            .TxtImp2 = nPorcentaje2
                            .TxtImp3 = nPorcentaje3
                        End With
                        lAgregaItem = False
                        FrmSeleccionProductos.lAgregaItem = False
                        FrmSeleccionProductos.Caption = "Modificar Artículos"
                        FrmSeleccionProductos.Adicionar(4).Enabled = True
                        'FrmSeleccionProductos.Flagloctem = lOC
                        FrmSeleccionProductos.Show vbModal
                        
                        CalculaTmp
                    End If
                End If
            Else                    ' Modifica en la Tabla
                If RsDocumentoD.RecordCount > 0 Then
                    If grdGrilla.VisibleRows > 0 Then
                        With FrmSeleccionProductos
                            .tEstadoDocu = "01" 'frmDocumento.RsCabecera!CabEstadoDocumento ' clr 15122011
                            .SubAsignar
                            .TxtCodigo.Enabled = False
                            '.TxtCodigoBarra.Enabled = False
                            .txtArticulo.Enabled = False
                            '.CmdCreaArticulo.Enabled = False
                            .ChkImp1.Caption = sImpuesto1
                            .ChkImp2.Caption = sImpuesto2
                            .ChkImp3.Caption = sImpuesto3
                            If pais <> "005" Then .TxtImp1 = nPorcentaje1
                            .TxtImp2 = nPorcentaje2
                            .TxtImp3 = nPorcentaje3
                        End With
                        lAgregaItem = False
                        FrmSeleccionProductos.lAgregaItem = False
                        FrmSeleccionProductos.Caption = "Modificar Artículos"
                        FrmSeleccionProductos.Adicionar(3).Enabled = False
                        FrmSeleccionProductos.Adicionar(4).Enabled = True
                        'FrmSeleccionProductos.Flagloctem = lOC
                        FrmSeleccionProductos.Show vbModal
                    End If
                End If
            End If
        Case Is = 2 'Eliminar
        
        
            lNavegar = False
            lLoad = False
            ValidaCierre
            If sCerrado = "Verdadero" Then Exit Sub
            
            If grdGrilla.VisibleRows = 0 Then
                'frmDocumentoDetalle.cmdOpcion(10).Enabled = False
                'frmDocumentoDetalle.cmdDescuento.Enabled = True      'False
                MsgBox "No existe datos para Eliminar...", vbInformation, sMensaje
                Exit Sub
            End If
        
            If RsDtemporal.RecordCount > 0 Then
                If MsgBox("¿Seguro de eliminar el Item " & Trim(grdGrilla.Columns(1).value) & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                    Exit Sub
                End If
                CnAlmacen.Execute "Delete From " & sDtemporal & " Where tCodigoProducto = '" & Trim(grdGrilla.Columns(0).value) & "'"
                RsDtemporal.Requery
                If RsDtemporal.RecordCount > 0 Then
                    RsDtemporal.MoveLast
                End If
                CalculaTmp
            End If
            
            LblItems = LblItems - 1
'            Select Case sPais
'                Case "001"
                    lCambioDetalle = True
'            End Select
        Case Is = 3 'Cancelar, blanquear formulario
            ValidaCierre
            If sCerrado = "Verdadero" Then Exit Sub
            
           If MsgBox("¿Seguro de cancelar el ingreso?", vbQuestion + vbOKCancel, sMensaje) = vbOK Then
                    '************** BLANQUEA LA CABECERA **************
                    CboTipoDocumento.BoundText = ""
                    LblPDoc = ""
                    TxtSerie = ""
                    TxtCorrelativo = ""
                    TxtProveedor = ""
                    TxtRUC = ""
                    TxtNombre = ""
                    CboArea.BoundText = ""
                    txtTransporte.Text = "0.00"
                    CboTipoIngreso.BoundText = ""
                    TxtGlosaDocumento = ""
                    dtcLibroCompra.Text = ""
                    CboMoneda.BoundText = ""
                    'CboReferencia.BoundText = ""
'                    TxtSerieRef = ""
'                    TxtCorrelativoRef = ""
                    TxtGlosaDocumento = ""
                    dtcLibroCompra.Text = ""
                    CboTipoDocumento.SetFocus
                    '********* BLANQUEA EL DETALLE (GRILLA) ***********
                    Set RsDtemporal = Nothing
                    CargaTemporal
                    LblItems = 0
            End If
    End Select
    
    If grdGrilla.VisibleRows = 0 Then
        frmDocumentoCompra.cmdOpcion(10).Enabled = False       'Botón de Revertir impuestos
        'frmDocumentoCompra.cmdDescuento.Enabled = True         'False
    Else
        'Bolivia
        Select Case pais
            Case "000", "004", "005": frmDocumentoCompra.cmdOpcion(10).Enabled = True     'Botón de Revertir impuestos
            Case "001": frmDocumentoCompra.cmdOpcion(10).Enabled = False
            Case "002": frmDocumentoCompra.cmdOpcion(10).Enabled = True       'Botón de Revertir impuestos
        End Select
        If LblSumDescuento = "" Then LblSumDescuento = 0
        If CmdProcesar.Enabled And CDbl(LblSumDescuento) > 0 Then
            'frmDocumentoCompra.cmdDescuento.Enabled = True
        End If
    End If
End Sub
Sub CalculaTmp()
    Select Case pais
        Case "000", "002", "004", "003", "005"
                If Not lNavegar And Not lLoad Then
                    Isql = "Select Sum(nImpuesto1) as nImpuesto1, Sum(nImpuesto2) as nImpuesto2, Sum(nImpuesto3) as nImpuesto3, Sum(nNeto) as nNeto, Sum(nDescuento) as nDescuento From " & sDtemporal & " "
                    Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
                    If Rs.RecordCount > 0 Then
                        wImp1 = IIf(IsNull(Rs!nImpuesto1), 0, Rs!nImpuesto1)
                        wImp2 = IIf(IsNull(Rs!nImpuesto2), 0, Rs!nImpuesto2)
                        wImp3 = IIf(IsNull(Rs!nImpuesto3), 0, Rs!nImpuesto3)
                        wSubTotal = IIf(IsNull(Rs!nNeto), 0, Rs!nNeto)
                        wDesc = IIf(IsNull(Rs!nDescuento), 0, Rs!nDescuento)
                    Else
                        wImp1 = 0
                        wImp2 = 0
                        wImp3 = 0
                        wSubTotal = 0
                        wDesc = 0
                    End If
        
                    LblTotImp1 = Format(wImp1, "##,##0.00")
                    LblTotImp2 = Format(wImp2, "##,##0.00")
                    LblTotImp3 = Format(wImp3, "##,##0.00")
                    LblSubTot = Format(wSubTotal, "##,##0.00")
'                    lblTotal = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo), "##,##0.00")
                    LblTotal = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(txtTransporte.Text), "##,##0.00")
'                    LblSaldo = LblTotal
'                    LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(lblTotLey1), "##,##0.00")
                    LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(lblTotLey1) + CDbl(txtTransporte.Text), "##,##0.00")
                    LblSumDescuento = Format(wDesc, "#,##0.00")
                End If
        Case "001" 'Bolivia
    If Not lNavegar And Not lLoad Then
        Isql = "Select Sum(nImpuesto1) as nImpuesto1, Sum(nImpuesto2) as nImpuesto2, Sum(nImpuesto3) as nImpuesto3, Sum(nNeto) as nNeto, Sum(nDescuento) as nDescuento From " & sDtemporal & " "
        Set Rs = Lib.OpenRecordset(Isql, Cn)
        If Rs.RecordCount > 0 Then
            wImp1 = IIf(IsNull(Rs!nImpuesto1), 0, Rs!nImpuesto1)
            wImp2 = IIf(IsNull(Rs!nImpuesto2), 0, Rs!nImpuesto2)
            wImp3 = IIf(IsNull(Rs!nImpuesto3), 0, Rs!nImpuesto3)
            wSubTotal = IIf(IsNull(Rs!nNeto), 0, Rs!nNeto)
            wDesc = IIf(IsNull(Rs!nDescuento), 0, Rs!nDescuento)
        Else
            wImp1 = 0
            wImp2 = 0
            wImp3 = 0
            wSubTotal = 0
            wDesc = 0
        End If
        
        LblTotImp1 = Format(wImp1, "##,##0.00")
        LblTotImp2 = Format(wImp2, "##,##0.00")
        LblTotImp3 = Format(wImp3, "##,##0.00")
        'modificado clr 1009211

        LblSubTot = Format(wSubTotal, "##,##0.00")
'                lblTotal = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo), "##,##0.00")
        LblTotal = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(txtTransporte.Text), "##,##0.00")
        'LblSaldo = LblTotal
'                LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(lblTotLey1), "##,##0.00")
        LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(lblTotLey1) + CDbl(txtTransporte.Text), "##,##0.00")
        LblSumDescuento = Format(wDesc, "#,##0.00")
    End If
    End Select
    
'    If lActualiza Then
'        LblSaldo = LblTotal
'        lActualiza = False
'    End If
End Sub
Private Sub cmdDoaceptar_Click()
   If swObliga Then
      MsgBox "Ingrese el motivo de la anulación...", vbExclamation, sMensaje
      txtDobserva.SetFocus
      Exit Sub
   Else
      FraDobserva.Visible = False
   End If
End Sub

Private Sub cmdDocancelar_Click()
   txtDobserva = ""
   txtDobserva.SetFocus
End Sub

Private Sub cmdDograbar_Click()
    sDocumento = Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo)

    If txtDobserva = "" Then
        MsgBox "Ingrese el motivo de la anulación...", vbExclamation, sMensaje
        txtDobserva.SetFocus
        Exit Sub
    End If
    
    Isql = "Update mDocumento Set " & _
           "tObservacion = '" & txtDobserva & " (" & sUsuario & " - " & Format(FechaServidor, "dd/mm/yyyy") & " " & Format(FechaServidor, "hh:mm AM/PM") & ")'" & " ,lreplica=1  where tCorrelativo = '" & LblCorrelativo & "'"
    CnAlmacen.Execute Isql
    Dim iRow As Integer
'    frmDocumento.RsCabecera.Requery
'    If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.MoveFirst
'    frmDocumento.RsCabecera.Find "tCorrelativo='" & LblCorrelativo & "'"
'    If frmDocumento.RsCabecera.EOF Then frmDocumento.RsCabecera.MoveLast
    FraDobserva.Visible = False
    ActivarNavegar True
End Sub

Private Sub cmdFind_Click()
    'xISQL = ""
    Isql = "Select tCodigoProveedor As Codigo, tRazonSocial as Descripcion,tIndicadorTributario from tProveedor where lActivo=1 order by tRazonSocial"
    'frmBusca.lEntra = True
   frmBusca.nPredeterm = 2
   frmBusca.TipoOperacion = "Almacen"
   Call ConfGrilla(3, frmBusca.grdGrilla, "codigo", 2, "Codigo", 1000, 0, 0, "", _
                                          "Descripcion", 2, "Descripcion", 4100, 0, 0, "", _
                                          "IndicadorTributario", 2, "tIndicadorTributario", 2400, 0, 0, "")
    
    
    
    frmBusca.Show vbModal
    If Sw Then
        'TxtDiasCredito = '<Calcular("Select nCredito As Codigo From tProveedor Where tCodigoProveedor='" & sCodigo & "'", CnAlmacen)
        'TxtRUC.Text = frmBusca.Rs!IndicadorTributario
        TxtRUC.Text = Calcular("Select tIndicadorTributario As Codigo From tProveedor Where tCodigoProveedor='" & sCodigo & "'", CnAlmacen)
        TxtProveedor.Text = sCodigo
        TxtNombre.Text = sDescrip
    Else
        'TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
    End If
    'If cmdOpcion(3).Visible Then Exit Sub
    GeneraMensaje
End Sub







Function MensajeCabecera() As Boolean
        If CboTipoDocumento.Enabled Then If CboTipoDocumento.BoundText = "" Then MsgBox "Ingrese el Documento", vbExclamation, sMensaje: Me.CboTipoDocumento.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(TxtSerie) = "" Then MsgBox "Ingrese el Nº de Serie...", vbExclamation, sMensaje: TxtSerie.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(TxtCorrelativo) = "" Then MsgBox "Ingrese el Correlativo...", vbExclamation, sMensaje: TxtCorrelativo.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(TxtProveedor) = "" Or Len(Trim(TxtProveedor)) <> 5 Then MsgBox "Ingrese el Proveedor...", vbExclamation, sMensaje:  MensajeCabecera = True: Exit Function
        If Trim(TxtRUC) = "" Then MsgBox "Ingrese el R.U.C...", vbExclamation, sMensaje: TxtRUC.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(CboArea.BoundText) = "" Then MsgBox "Ingrese el Área...", vbExclamation, sMensaje: CboArea.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(CboTipoIngreso.BoundText) = "" Then MsgBox "Ingrese el Tipo de Ingreso...", vbExclamation, sMensaje: CboTipoIngreso.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(TxtGlosaDocumento) = "" Then MsgBox "Ingrese la Glosa...", vbExclamation, sMensaje: TxtGlosaDocumento.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(CboMoneda.BoundText) = "" Then MsgBox "Ingrese la Moneda...", vbExclamation, sMensaje: Me.CboMoneda.SetFocus: MensajeCabecera = True: Exit Function
        If Trim(CboMoneda.BoundText) = "02" And CDbl(LblTipoCambio) <= 0 Then MsgBox "Ingrese el Tipo de Cambio en almacen...", vbExclamation, sMensaje: MensajeCabecera = True: Exit Function
        'If lProgramaPagoPre And dtcProgramaPago = "" Then MsgBox "Ingrese el Programa de Pagos...", vbExclamation, sMensaje: dtcProgramaPago.SetFocus: MensajeCabecera = True: Exit Function
        'If lLibroCompra And dtcLibroCompra = "" Then MsgBox "Ingrese el Libro de Compra...", vbExclamation, sMensaje: dtcLibroCompra.SetFocus: MensajeCabecera = True: Exit Function
 
End Function



Private Sub cmdObservacion_Click()
'    With frmDocumento.RsCabecera
'        If Not .EOF Then
'            If !CabEstadoDocumento = "04" Then    'Anulado
'                swObliga = False
'                txtDobserva = IIf(Not frmDocumento.RsCabecera.EOF, IIf(Not IsNull(frmDocumento.RsCabecera!CabObservacion), frmDocumento.RsCabecera!CabObservacion, ""), "")
'                If txtDobserva = "" Then
'                    LblAviso(4).Caption = "Ingrese el Motivo de la Anulación"
'                    cmdDograbar.Visible = True
'                    cmdDocancelar.Visible = True
'                    cmdDoaceptar.Left = 3240
'                Else
'                    LblAviso(4).Caption = "Observación de la Anulación"
'                    cmdDograbar.Visible = False
'                    cmdDocancelar.Visible = False
'                    cmdDoaceptar.Left = cmdDocancelar.Left
'                End If
'                txtDobserva.MaxLength = 120
'                FraDobserva.Visible = True
'            End If
'        End If
'    End With
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
        Case Is = 0 'Agregar
            Dim lAgregaN As Boolean
            lAgregaN = False
            If lDiaContable = False Then 'manual
                If lDiaContableAperturado = True Then
                    lAgregaN = True
                Else
                    'frmDiaContable.obtieneModoIngreso "Apertura"
                    'frmDiaContable.Show vbModal
                    lAgregaN = True
                End If
            Else ' automatico
                lAgregaN = True
            End If
            
            If lAgregaN Then
                Sw = True
                           
                sFlag = True
                
                Isql = "Select [fFecha], [nCompra], [nVenta] From tTipoCambio " _
                       & " Where fFecha='" & Format(FechaServidor, "MM/dd/yyyy") & "'"
                Set RsCambio = Lib.OpenRecordset(Isql, CnAlmacen)
                If Not RsCambio.EOF Then
                    RsCambio.MoveLast
                    LblTipoCambio.Caption = Format(RsCambio!nVenta, "##0.000")
                Else
                    LblTipoCambio.Caption = "0.000"
                End If
                
                LblEstado = ""
                LblUsuario = ""
                LblCorrelativo = ""
                LblPDoc = ""
                Blanquear Me
                '|||Jesus|||
                'LblPDocRef = ""
                lCanje = False
                '|||||||||||
                cmdOpcion(3).Enabled = False
                cmdOpcion(3).Visible = False
                cmdOpcion(2).Enabled = True
                cmdOpcion(2).Visible = True
                
                cmdOpcion(4).Enabled = True
                'cmdDescuento.Enabled = True     'False
                CmdProcesar.Enabled = False
                ActivaAdicion True
                ActivarBotones False
                DtpEmision.value = Format(FechaServidor, "dd/mm/yyyy")
                DtpIngreso.value = Format(FechaServidor, "dd/mm/yyyy")
                'DtpFPrograma.value = Format(FechaServidor, "dd/mm/yyyy")
                dtpAnio.value = Format(FechaServidor, "dd/mm/yyyy")
                EstadoCabecera True
                'ChkPrograma.Enabled = True
                'dtcProgramaPago.Enabled = True
                LlenaMesContable
                LlenaGrilla
                'SwTmp = True
                CargaTemporal
                CboArea.BoundText = "000"
                txtTransporte.Text = "0.00"
                CboMoneda.BoundText = "01"
                CboTipoIngreso.BoundText = "M"
                cboOperacion.BoundText = Calcular("select tCodigo as Codigo from tTabla where tValor='G' and tTabla='TOPERACION' and lActivo=1", CnAlmacen)
                cboOperacion.Enabled = False
                'CboReferencia.Enabled = False
                
                LblSubTot = "0.00"
                LblTotImp1 = "0.00"
                LblTotImp2 = "0.00"
                LblTotImp3 = "0.00"
                LblTotal = "0.00"
                txtMontoControl = "0.00"
                LblSubTot = "0.00"
                LblSumDescuento = "0.00"
                LblSaldo = "0.00"
                lblRedondeo = "0.00"
                lblTotLey1 = "0.00"
                lblGranTotal = "0.00"
                
                LblTotImp1.ToolTipText = ""
                LblTotImp2.ToolTipText = ""
                LblTotImp3.ToolTipText = ""
                lblRedondeo.ToolTipText = ""
                LblTotal.ToolTipText = ""
                

                'ChkPrograma = 1
                DtpIngreso.Enabled = True
                            
                
            End If
        Case Is = 2  'Anula Documento
  
        Case Is = 4  'Grabar
            Dim sCorrela     As String
            Dim sCorrelativo As String
            Dim nPos         As Integer
            
            sFlag = False
            If MensajeCabecera Then
                'FinModifica
                Exit Sub
            End If

            Screen.MousePointer = vbHourglass
            If grdGrilla.VisibleRows = 0 Then
                MsgBox "No existe datos a procesar!!!", vbInformation, sMensaje
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            ValidaCierre
            If sCerrado = "Verdadero" Then
                'FinModifica
                Exit Sub
            End If
            
            ValidaPlazo

            
            If wVaRegCompra Then
                Select Case pais
                    Case "001"
'                        If txtCodigoControl.Text = "" Then
'                            MsgBox "Debe especificar el Código de Control.", vbCritical, sMensaje
'                            Screen.MousePointer = vbDefault
'                            txtCodigoControl.SetFocus
'                            Exit Sub
'                        End If
'                        If txtAutorizacion.Text = "" Then
'                            MsgBox "Debe especificar el Código de Autorización.", vbCritical, sMensaje
'                            Screen.MousePointer = vbDefault
'                            txtAutorizacion.SetFocus
'                            Exit Sub
'                        End If
                    Case "002"

'                        If Format(DtpIngreso.value, "dd/mm/yyyy") = Format(dtpFechCaducidad.value, "dd/mm/yyyy") Then
'                            If MsgBox("La Fecha de Caducidad es igual a la fecha de Ingreso, ¿Desea continuar?...", vbQuestion + vbYesNo, sMensaje) = vbNo Then
'                                Screen.MousePointer = vbDefault
'                                Exit Sub
'                            End If
'                        End If

                    

                End Select
            End If
                        
            nLey1 = FuncPercepcion(Sw, LblCorrelativo)
            
            'JESUS
            'If FrmDocumento.RsCabecera!CabEstadoDocumento = "01" Then
            nLeyAD = FuncPorcentajeLeyAD(Sw, LblCorrelativo)
            'End If
            '------02.11.2017--------
            
            If Sw Then 'Nuevo Documento
                If VerificarExistenciaDocumento = False Then Screen.MousePointer = vbDefault: Exit Sub
                If CboArea.BoundText = "" Then
                    MsgBox "se tiene que Elegir Area", vbCritical, sMensaje
                    CboArea.SetFocus
                    Exit Sub
                End If
                
                xAnho = Year(DtpIngreso)
                xMes = Mid("00", 1, 2 - Len(Trim(str(Month(DtpIngreso))))) & Trim(str(Month(DtpIngreso)))
                xPeriodo = xAnho & xMes
                nCorrela = Calcular("Select Right(Max(tCorrelativo),4) As Codigo From MDocumento Where Left(tCorrelativo,6)='" & xPeriodo & "'", CnAlmacen)
                
                If IsNull(nCorrela) Or nCorrela = "0" Or nCorrela = "" Then
                    LblCorrelativo = xPeriodo & "-" & "0001"
                Else
                    LblCorrelativo = xPeriodo & "-" & Lib.Correlativo(nCorrela, 4)
                End If
                
                nTotal = CDbl(LblTotal)
                nRedondeo = CDbl(lblRedondeo)
                LblSaldo = Format(nTotal + nRedondeo, "###,###,##0.00") 'punto 1
                nGranTotal = CDbl(LblTotal) + nLey1
                lblGranTotal = Format(nGranTotal, "###,###,##0.00")
            Else 'Modifica Documento
                If VerificarExistenciaDocumento = False Then Screen.MousePointer = vbDefault: Exit Sub
                lblTotLey1 = Format(nLey1, "###,###,##0.00")
                LblSaldo = Format((CDbl(LblTotal) + CDbl(lblTotLey1)) - VerificaAbonosNC(TxtProveedor, Trim(LblPDoc) + Trim(TxtSerie) + Trim(TxtCorrelativo)), "###,###,##0.00")
                nGranTotal = CDbl(LblTotal) + nLey1
                lblGranTotal = Format(nGranTotal, "###,###,##0.00")
                If Not cmdOpcion(3).Visible Then
                    If iTBox = 0 Then ' hay error no pasa por aca
                        lActualizaImpuesto = True
                        ActDetalle
                        lActualizaImpuesto = False
                    Else
                        LlenaGrilla
                    End If
                Else
                    cmdOpcion(3).Enabled = True
                    'If iTBox = 0 Then ModificaDocumento
                End If
                lModificado = True
            End If
                        
            '********PROCESO DE GUARDADO************
            Screen.MousePointer = vbHourglass
            
            Set oDiaContable = New clsDiaContable
            Set clsmDocumento = New ClsDocumento
            

            
            sCorrecto = clsmDocumento.InsUpdmDocumento(Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo), _
                                                        Format(DtpEmision, "yyyy/mm/dd"), _
                                                        Format(DtpIngreso, "yyyy/mm/dd"), _
                                                        CDbl(LblTotal), CDbl(LblSumDescuento), CDbl(LblSubTot), _
                                                        CDbl(LblTotImp1), CDbl(LblTotImp2), CDbl(LblTotImp3), _
                                                        TxtProveedor, "", sUsuario, "01", "", _
                                                        LblCorrelativo, CDbl(LblTipoCambio.Caption), TxtGlosaDocumento, CboTipoIngreso.BoundText, _
                                                        CboMoneda.BoundText, CboArea.BoundText, "", "", "", "", "", _
                                                        CDbl(LblSaldo), Trim(CboTipoDocumento.BoundText), "", lCanje, IIf(lDesc = "", "Null", IIf(lDesc = "G", "1", "0")), _
                                                        CDbl(lblRedondeo), cboOperacion.BoundText, False, nLey1, "", "", LblICE.Caption, "", sUsuarioAutoriza, "", "", "", CDbl(txtTransporte.Text), dtcLibroCompra.BoundText, Year(dtpAnio.value), CmbMes.ListIndex + 1, IIf(Sw, "I", "U"))

           
            If sCorrecto = "X" Then
                MsgBox "Ocurrio un error al tratar de insertar/Modificar los datos del Documento", vbInformation, sMensaje
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            If Sw Then
                ActualizaDetalle
                Actualiza
                RsDtemporal.Requery
            Else
                RsDocumentoD.Requery
            End If
          
            
            Screen.MousePointer = vbDefault
'            If Sw Then
'                MsgBox "Registro guardado...", vbInformation, sMensaje
'            Else
'                MsgBox "Registro modificado...", vbInformation, sMensaje
'            End If
            Sw = False
            'SwTmp = False
            
            '********FIN DE PROCESO DE GUARDADO************
                        
            If iTBox = 0 Then
                swMover = True
                swEntrar = False
                'Asignar
                swEntrar = True
                LlenaGrilla
                swMover = False
                ActivarBotones True
                ActivaAdicion True
                
                'FinModifica
            End If
            
            If Sw = False Then
                    LblTotImp1.ToolTipText = ""
                    LblTotImp2.ToolTipText = ""
                    LblTotImp3.ToolTipText = ""
                    lblRedondeo.ToolTipText = ""
                    LblTotal.ToolTipText = ""
            End If
            sUsuarioAutoriza = ""
            lCambioDetalle = False
            
            
            CmdProcesar_Click
            cmdPago_Click
            
            frmReciboEgresoDetalle.txtDescripcion.Text = TxtGlosaDocumento.Text ' "Compra mercaderia documento: " & LblPDoc.Caption & TxtSerie & TxtCorrelativo
            frmReciboEgresoDetalle.txtAutorizacion.Text = ""
            frmReciboEgresoDetalle.txtMonto.Text = lblGranTotal.Caption
            frmReciboEgresoDetalle.nMonto = CDbl(lblGranTotal.Caption)
            frmReciboEgresoDetalle.sMonedaRecibo = CboMoneda.BoundText
            'frmReciboEgresoDetalle.cboTipoEgreso.BoundText = "001"
            
            frmReciboEgresoDetalle.txtComprobante = Trim(LblPDoc.Caption) & Trim(TxtSerie.Text) & Trim(TxtCorrelativo.Text)
            frmReciboEgresoDetalle.txtnotacredito = LblCorrelativo.Caption
            frmReciboEgresoDetalle.txtIdentificacion = TxtRUC.Text
            frmReciboEgresoDetalle.txtRSocial = TxtNombre.Text
            frmReciboEgresoDetalle.txtTipoComprobante = Me.CboTipoDocumento.Text
            frmReciboEgresoDetalle.sCompraOK = "OK"
            'frmReciboEgresoDetalle.cmdOpcion(1)
'              nMonto = IIf(wEnter, val(sDescrip), nMonto)
'               txtMonto.Text = Format(nMonto, "###,##0.00")
'               cmdOpcion(6).FontBold = True
'               cmdOpcion(7).FontBold = False
'               cmdOpcion(9).FontBold = False
'               cmdOpcion(10).FontBold = False
'               sMonedaRecibo = "01"
            
            
            Unload Me
        Case Is = 3 'Modificar
'            Dim fProce As Date
'            fProce = Format(txtFechaProceso.Text, "dd/mm/yyyy")
'
'            ValidaCierre
'            If sCerrado = "Verdadero" Then Exit Sub
'
'            ValidaPlazo
'
'
'            'Modificación CLR 11/01/2012 comprobar que no se cambie el valor de compra de documentos cancelados
'            If lModificarDocumento = False And frmDocumento.RsCabecera!CabEstadoDocumento = "02" Then
'                MsgBox "El documento ya fue procesado. " & vbCr & "Si desea modificarlo debe desprocesarlo antes.", vbInformation, sMensaje
'                Exit Sub
'            End If
'            If lModificarDocumento = False And frmDocumento.RsCabecera!CabEstadoDocumento = "03" Then
'                MsgBox "El documento ya fue Cancelado. " & vbCr & "Si desea modificarlo debe quitar el pago de cancelación y " & vbCr & "luego desprocesarlo.", vbInformation, sMensaje
'                Exit Sub
'            End If
'            'fin modificación CLR 11/01/2012
'
'            If Contraseña("01") = False Then
'                MsgBox "Clave no permitida", vbExclamation, sMensaje
'                Exit Sub
'            End If
'            'Modificación CLR 15/12/2011 comprobar que no se cambie el valor de compra de documentos cancelados
'
'            If frmDocumento.RsCabecera!CabEstadoDocumento = "03" Then
'                If MsgBox("El documento está cancelado. " & vbCr & "Si desea cambiar el valor del total del documento," & vbCr & "antes debe quitar los pagos de cancelación." & vbCr & "¿Desea modificar el valor del total del documento?", vbQuestion + vbYesNoCancel, sMensaje) <> vbYes Then
'                    MsgBox "Modificación Cancelada", vbExclamation + vbOKOnly, sMensaje
'                    Exit Sub
'                End If
'            End If
'
'            'Fin Modificación CLR 15/12/2011
'            cmdOpcion(4).Enabled = True 'Habilita el boton grabar
'            cmdOpcion(3).Enabled = False 'Desabilita el boton modificar
'
'            'Jesus 08-02-16
'            'Se agrego validacion de estado para modificacion del detalle
'            If ValidaDuplicidad Then
'                'Jesus 08-07-2016  ** validacion por base de datos
'                lModificarDetalle = Calcular("select IsNull(lModificarDetalleDoc,0) as Codigo from tParametro ", CnAlmacen)
'                If lModificarDetalle Then
'                    BotonesDetalle True
'                Else
'                    BotonesDetalle False
'                End If
'                'Fin cambio
'            Else
'                BotonesDetalle True
'            End If
'            'Fin cambio
'
'            CapturaDatos
'            EstadoCabecera True
'
'            CboArea.Enabled = False
''            txtTransporte.Enabled = False
'            CboReferencia.Enabled = False
'            LblPDocRef.Enabled = False
'            TxtSerieRef.Enabled = False
'            TxtCorrelativoRef.Enabled = False
'
'            DtpEmision.Enabled = True
'            DtpIngreso.Enabled = True
'            CboTipoIngreso.Enabled = False
'            CboMoneda.Enabled = False
'
'            ChkPrograma.Enabled = True
'            dtcProgramaPago.Enabled = True
'            If ChkPrograma.value Then
'                    DtpFPrograma.Enabled = True
'                    TxtDiasCredito.Enabled = True
'            Else
'                    DtpFPrograma.Enabled = False
'                    TxtDiasCredito.Enabled = False
'            End If
'
'            CboTipoDocumento.SetFocus
'
'            If ErpOF_Activo Then
'                CboTipoDocumento.Enabled = False
'                TxtSerie.SetFocus
'            End If
    
        Case Is = 6 'O C         'Cancelar
            
    
        Case Is = 7 'Canje
            
         
        Case Is = 5 'Emite
            
           
        Case Is = 8 'Historial de Pagos
            
        Case Is = 9 'Salir
            Dim nSald As Double
            Dim nTota As Double
            FraDobserva.Visible = False
            
            Unload Me
            
        Case Is = 10        'REVERTIR IMPUESTOS ************
           If lDesc = "I" Or lDesc = "G" Then
                MsgBox "Se a detectado que el documento tiene descuentos aplicados, " & vbCr & "no es posible cambiar los impuestos.", vbInformation + vbOKOnly, sMensaje
                Exit Sub
            End If
          
            If grdGrilla.VisibleRows = 0 Then
                MsgBox "No existen datos. El ajuste de impuestos no procede...", vbInformation, sMensaje
                Exit Sub
            End If
         
            ValidaCierre
            If sCerrado = "Verdadero" Then Exit Sub
        
            ValidaPlazo
            'if Not swPasar Then Exit Sub
      
            Dim xImp1, xImp2, xImp3 As String
            Dim pImp1, pImp2, pImp3, sTotPorcentaje, lImp1, lImp2, lImp3 As String
            Dim nTotImp, sTotal As Double
            Dim sSubTot As Double
            Dim nOtrosCargosInafecto As Double
         
            lActualiza = True
            Dim SwTmp As Boolean
            SwTmp = True
            If SwTmp Then           ' Si se está trabajando en el temporal...
                If RsDtemporal.RecordCount = 0 Then
                    MsgBox "No existe datos para Procesar...", vbExclamation, sMensaje
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                If RsDtemporal.RecordCount > 0 Then RsDtemporal.MoveFirst
                Do While Not RsDtemporal.EOF
                    sCodi = IIf(IsNull(RsDtemporal!tCodigoProducto), 0, RsDtemporal!tCodigoProducto)
                    sNeto = Round(IIf(IsNull(RsDtemporal!nNeto), 0, RsDtemporal!nNeto) - IIf(IsNull(RsDtemporal!nOtrosCargosInafecto), 0, RsDtemporal!nOtrosCargosInafecto), 6)
                    sDesc = IIf(IsNull(RsDtemporal!nDescuento), 0, RsDtemporal!nDescuento)
                    sReca = IIf(IsNull(RsDtemporal!nRecargo), 0, RsDtemporal!nRecargo)
                    sPVen = IIf(IsNull(RsDtemporal!nPVenta), 0, RsDtemporal!nPVenta) - IIf(IsNull(RsDtemporal!nOtrosCargosInafecto), 0, RsDtemporal!nOtrosCargosInafecto)
                    sPuni = IIf(IsNull(RsDtemporal!nPrecio), 0, RsDtemporal!nPrecio)
                    sCant = IIf(IsNull(RsDtemporal!nCantidad), 0, RsDtemporal!nCantidad)
                    xImp1 = IIf(IsNull(RsDtemporal!nImpuesto1), 0, RsDtemporal!nImpuesto1)
                    xImp2 = IIf(IsNull(RsDtemporal!nImpuesto2), 0, RsDtemporal!nImpuesto2)
                    xImp3 = IIf(IsNull(RsDtemporal!nImpuesto3), 0, RsDtemporal!nImpuesto3)
                    nOtrosCargosInafecto = IIf(IsNull(RsDtemporal!nOtrosCargosInafecto), 0, RsDtemporal!nOtrosCargosInafecto)
                    If sPVen <> "" Then
                        pImp1 = IIf(RsDtemporal!nPorcentaje1 = "", 0, RsDtemporal!nPorcentaje1)
                        pImp2 = IIf(RsDtemporal!nPorcentaje2 = "", 0, RsDtemporal!nPorcentaje2)
                        pImp3 = IIf(RsDtemporal!nPorcentaje3 = "", 0, RsDtemporal!nPorcentaje3)
                        nTotImp = CDbl(pImp1) + CDbl(pImp2) + CDbl(pImp3)
                        sTotPorcentaje = nTotImp / 100 + 1
                        If sPVen = sNeto + sDesc - sReca Then
                            sSubTot = sNeto / CDbl(sTotPorcentaje)
                        Else
                            sSubTot = sPVen - sDesc + sReca
                        End If
                        lImp1 = sSubTot * (CDbl(pImp1) / 100)
                        lImp2 = sSubTot * (CDbl(pImp2) / 100)
                        lImp3 = sSubTot * (CDbl(pImp3) / 100)
                        sSubTot = sSubTot + nOtrosCargosInafecto
                        sTotal = sSubTot + lImp1 + lImp2 + lImp3
                    End If
                    If sPVen <> "" And sCant <> "" Then
                        sPuni = CDbl(sSubTot / sCant)
                    Else
                        Exit Sub
                    End If
                 
                    Isql = "Update " & frmDocumentoCompra.sDtemporal & " Set " _
                           & " nTotal =" & CDbl(sTotal) & ", " _
                           & " nPrecio=" & CDbl(sPuni) & ", " _
                           & " nImpuesto1 =" & IIf(CDbl(lImp1) <> 0, CDbl(lImp1), 0) & ", " _
                           & " nImpuesto2 =" & IIf(CDbl(lImp2) <> 0, CDbl(lImp2), 0) & ", " _
                           & " nImpuesto3 =" & IIf(CDbl(lImp3) <> 0, CDbl(lImp3), 0) & ", " _
                           & " nNeto=" & CDbl(sSubTot) & ", " _
                           & " nDescuento=" & CDbl(sDesc) & ", " _
                           & " nRecargo=" & CDbl(sReca) & "" _
                           & " Where (tCodigoProducto='" & sCodi & "' And tCorrelativo='" & LblCorrelativo & "')" _
                           & " And nImpuesto1 =" & CDbl(xImp1) & "" _
                           & " And nImpuesto2 =" & CDbl(xImp2) & "" _
                           & " And nImpuesto3 =" & CDbl(xImp3) & ""
                    CnAlmacen.Execute Isql
                    RsDtemporal.MoveNext
                Loop
                
                frmDocumentoCompra.RsDtemporal.Requery
                frmDocumentoCompra.RsDtemporal.MoveFirst
                Screen.MousePointer = vbDefault
                
            Else                    ' Si se está trabajando en la tabla...
              ' A veces no se establece la variable revisar
                If RsDocumentoD.RecordCount > 0 Then RsDocumentoD.MoveFirst
                Screen.MousePointer = vbHourglass
                Do While Not RsDocumentoD.EOF
                    sCodi = IIf(IsNull(RsDocumentoD!tCodigoProducto), 0, RsDocumentoD!tCodigoProducto)
                    sNeto = Round(IIf(IsNull(RsDocumentoD!nNeto), 0, RsDocumentoD!nNeto) - IIf(IsNull(RsDocumentoD!nOtrosCargosInafecto), 0, RsDocumentoD!nOtrosCargosInafecto), 6)
                    sDesc = IIf(IsNull(RsDocumentoD!nDescuento), 0, RsDocumentoD!nDescuento)
                    sReca = IIf(IsNull(RsDocumentoD!nRecargo), 0, RsDocumentoD!nRecargo)
                    sPVen = IIf(IsNull(RsDocumentoD!nPVenta), 0, RsDocumentoD!nPVenta) - IIf(IsNull(RsDocumentoD!nOtrosCargosInafecto), 0, RsDocumentoD!nOtrosCargosInafecto)
                    sPuni = IIf(IsNull(RsDocumentoD!nPrecio), 0, RsDocumentoD!nPrecio)
                    sCant = IIf(IsNull(RsDocumentoD!nCantidad), 0, RsDocumentoD!nCantidad)
                    xImp1 = IIf(IsNull(RsDocumentoD!nImpuesto1), 0, RsDocumentoD!nImpuesto1)
                    xImp2 = IIf(IsNull(RsDocumentoD!nImpuesto2), 0, RsDocumentoD!nImpuesto2)
                    xImp3 = IIf(IsNull(RsDocumentoD!nImpuesto3), 0, RsDocumentoD!nImpuesto3)
                    nOtrosCargosInafecto = IIf(IsNull(RsDocumentoD!nOtrosCargosInafecto), 0, RsDocumentoD!nOtrosCargosInafecto)
                    If sPVen <> "" Then
                        pImp1 = IIf(RsDocumentoD!nPorcentaje1 = "", 0, RsDocumentoD!nPorcentaje1)
                        pImp2 = IIf(RsDocumentoD!nPorcentaje2 = "", 0, RsDocumentoD!nPorcentaje2)
                        pImp3 = IIf(RsDocumentoD!nPorcentaje3 = "", 0, RsDocumentoD!nPorcentaje3)
                        nTotImp = CDbl(pImp1) + CDbl(pImp2) + CDbl(pImp3)
                        sTotPorcentaje = nTotImp / 100 + 1
                             
                        Select Case pais
                            Case "000", "002", "003", "005"
                               
                                If val(sPVen) = val(sNeto) + val(sDesc) - val(sReca) Then
                                    sSubTot = sNeto / CDbl(sTotPorcentaje)
                                Else
                                    sSubTot = sPVen - sDesc + sReca
                                End If
                                     
                                lImp1 = sSubTot * (CDbl(pImp1) / 100)
                                lImp2 = sSubTot * (CDbl(pImp2) / 100)
                                lImp3 = sSubTot * (CDbl(pImp3) / 100)
                                sSubTot = sSubTot + nOtrosCargosInafecto
                                sTotal = sSubTot + lImp1 + lImp2 + lImp3
                                
                            Case "001"
                                sTotal = sNeto
                                If sPVen = sNeto + CDbl(sDesc) - CDbl(sReca) Then
                                    sSubTot = sNeto - (xImp1 + xImp2 + xImp3)
                                Else
                                    sSubTot = sPVen - sDesc + sReca
                                End If
                                lImp1 = sTotal * (CDbl(pImp1) / 100)
                                lImp2 = sTotal * (CDbl(pImp2) / 100)
                                lImp3 = sTotal * (CDbl(pImp3) / 100)
                                
                        End Select
                                                     
                    End If
                    If sPVen <> "" And sCant <> "" Then
                        sPuni = CDbl(sSubTot / sCant)
                    Else
                        Exit Sub
                    End If
                      
                    Isql = "Update dDocumento Set" _
                           & " nTotal =" & CDbl(sTotal) & ", " _
                           & " nPrecio=" & CDbl(sPuni) & ", " _
                           & " nImpuesto1 =" & IIf(CDbl(lImp1) <> 0, CDbl(lImp1), 0) & ", " _
                           & " nImpuesto2 =" & IIf(CDbl(lImp2) <> 0, CDbl(lImp2), 0) & ", " _
                           & " nImpuesto3 =" & IIf(CDbl(lImp3) <> 0, CDbl(lImp3), 0) & ", " _
                           & " nNeto=" & CDbl(sSubTot) & ", " _
                           & " nDescuento=" & CDbl(sDesc) & ", " _
                           & " nRecargo=" & CDbl(sReca) & "" _
                           & " Where (tCodigoProducto='" & sCodi & "' And tCorrelativo='" & LblCorrelativo & "')" _
                           & " And nImpuesto1 =" & CDbl(xImp1) & "" _
                           & " And nImpuesto2 =" & CDbl(xImp2) & "" _
                           & " And nImpuesto3 =" & CDbl(xImp3) & ""
                    Cn.Execute Isql
                    RsDocumentoD.MoveNext
                Loop
                lLoad = False
                lNavegar = False
                CalculaTabla
                Actualiza
                lCambioDetalle = True
'
                frmDocumentoCompra.RsDocumentoD.Requery
                'iRow = frmDocumento.RsCabecera.AbsolutePosition
                'frmDocumento.RsCabecera.Requery
                'If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.AbsolutePosition = IIf(iRow = -1, 1, iRow)
                
                If frmDocumentoCompra.RsDocumentoD.RecordCount > 0 Then
                    frmDocumentoCompra.RsDocumentoD.MoveFirst
                End If
                
                Screen.MousePointer = vbDefault
            End If
            lNavegar = False
          
          CalculaTmp
          
    End Select
    If lModificado = False Then
        If (lDatosExternos = True And Me.lGrabo = True) Or lDatosExternos = False Then
            If Index <> 9 Then
                If LblCorrelativo <> "" Then
                    lLoad = False
                    CalculaTabla
                    lLoad = True
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
    End If
    lDatosExternos = False
    lModificado = False
End Sub

Sub Actualiza()
    Isql = "Update mDocumento Set" _
           & " nNeto=" & CDbl(LblSubTot) & ", " _
           & " nTotal=" & CDbl(LblTotal) & ", " _
           & " nImpuesto1=" & CDbl(LblTotImp1) & ", " _
           & " nImpuesto2=" & CDbl(LblTotImp2) & ", " _
           & " nImpuesto3=" & CDbl(LblTotImp3) & ", " _
           & " nSaldo=" & CDbl(LblSaldo) & ", " _
           & " nDescuento=" & CDbl(LblSumDescuento) & ", " _
           & " lDescuento=" & IIf(lDesc = "", "Null", IIf(lDesc = "G", 1, 0)) & "" _
           & " ,lreplica=1 Where tcorrelativo='" & LblCorrelativo & "'"
    CnAlmacen.Execute Isql
End Sub



Private Function ValidaPrecioPromedio() As String
Dim nPorcFN As Double
Dim nPorcFP As Double
Dim nPorcFR As Double
Dim sPorcFR As String
Dim nValPorcFN As Double
Dim nValPorcFP As Double
Dim sFueraFluct As String
Dim i As Integer
Dim lConsoli As Boolean

    'Jesus 16/03/16
    sTempoInsumos = dbTemporal(sUsuario, 7, "tCodigo", "nvarchar(20) COLLATE Modern_Spanish_CI_AS NULL ", _
                                            "tArticulo", "nvarchar(50) COLLATE Modern_Spanish_CI_AS NULL ", _
                                            "nPrecioPromedio", "Float", _
                                            "nFlucPos", "nvarchar(50) COLLATE Modern_Spanish_CI_AS NULL ", _
                                            "nFlucNeg", "nvarchar(50) COLLATE Modern_Spanish_CI_AS NULL ", _
                                            "nPrecioPromedioNuevo", "Float", _
                                            "nFlugRes", "nvarchar(50) COLLATE Modern_Spanish_CI_AS NULL ")
    'Fin 16/03/16
    
    If Sw Then
        sFueraFluct = ""
        RsDtemporal.MoveFirst
        For i = 0 To RsDtemporal.RecordCount - 1
        
'            lConsoli = Calcular("SELECT isnull(lPPConsolidadoSA,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDtemporal!tCodigoProducto & "')", CnAlmacen)
'            If cboMoneda.BoundText = "01" Then      'Moneda Nacional
'                If lPPConsolidado Then
'                    If lConsoli Then
'                        Isql = "SELECT a.tCodigoProducto, a.nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                               & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + a.nCantidad * a.nPrecio) " _
'                               & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                               & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact,b.tDetallado Producto " _
'                               & " FROM dbo.DDOCUMENTO AS a INNER JOIN dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto left JOIN " _
'                               & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                               & " WHERE (b.tCodigoProducto = '" & RsDtemporal!tCodigoProducto & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'                    Else
'                        Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                               & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * a.nprecio))" _
'                               & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact,b.tDetallado Producto From ddocumento a," _
'                               & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDtemporal!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                    End If
'                Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * a.nprecio))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDtemporal!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                End If
'            Else                                    'Moneda Extranjera
'                If lPPConsolidado Then
'                    If lConsoli Then
'                        Isql = "SELECT a.tCodigoProducto, a.nPrecio*c.ncambio as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                               & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + a.nCantidad * (a.nPrecio * c.nCambio)) " _
'                               & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                               & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact,b.tDetallado Producto " _
'                               & " FROM dbo.DDOCUMENTO AS a INNER JOIN " _
'                               & " dbo.MDOCUMENTO AS c ON a.tCorrelativo = c.tCorrelativo INNER JOIN " _
'                               & " dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto LEFT JOIN " _
'                               & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                               & " WHERE (b.tCodigoProducto = '" & RsDtemporal!tCodigoProducto & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'
'                    Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio*c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (a.nPrecio * c.nCambio)))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDtemporal!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'
'                    End If
'                Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio*c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (a.nPrecio * c.nCambio)))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDtemporal!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                End If
'            End If
            
            Isql = "exec usp_Almacen_CalculaPrecioPromedio '" & LblCorrelativo & "', '" & RsDtemporal!tCodigoProducto & "', ''"
        
            Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)

            If RsPromedio.RecordCount > 0 Then
                RsPromedio.MoveFirst
                nPorcFN = Calcular("SELECT isnull(nPorcFlucNeg,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDtemporal!tCodigoProducto & "')", CnAlmacen)

                nPorcFP = Calcular("SELECT isnull(nPorcFlucPos,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDtemporal!tCodigoProducto & "')", CnAlmacen)

                'RsPromedio!nPrecioPromedio
                If nPorcFN <> 0 Or nPorcFP <> 0 Then
                    nValPorcFN = RsPromedio!nPrecioPromedio - (nPorcFN * RsPromedio!nPrecioPromedio / 100)
                    nValPorcFP = RsPromedio!nPrecioPromedio + (nPorcFP * RsPromedio!nPrecioPromedio / 100)
                    If Not (nValPorcFN <= RsPromedio!PrecioPromedio And RsPromedio!PrecioPromedio <= nValPorcFP) And (nValPorcFN > 0 And nValPorcFP > 0) Then
                        sFueraFluct = sFueraFluct & vbCr & RsPromedio!tCodigoProducto & " - " & RsPromedio!Producto
                        
                        'Jesus 16/03/16
                        nPorcFR = ((RsPromedio!PrecioPromedio - RsPromedio!nPrecioPromedio) * 100 / RsPromedio!nPrecioPromedio)
                        If nPorcFR < 0 Then
                            sPorcFR = "(- " & Round(Abs(nPorcFR), 2) & "%)"
                        Else
                            sPorcFR = "(+ " & Round(Abs(nPorcFR), 2) & "%)"
                        End If
                        
                        Isql = "insert into " & sTempoInsumos & " (tCodigo,tArticulo,nPrecioPromedio,nFlucNeg,nFlucPos,nPrecioPromedioNuevo,nFlugRes) values " & _
                                "('" & RsPromedio!tCodigoProducto & "','" & RsPromedio!Producto & "', " & RsPromedio!nPrecioPromedio & ", " & _
                                " '(- " & nPorcFN & "%) : " & Round(nValPorcFN, 3) & "', '(+ " & nPorcFP & "%) : " & Round(nValPorcFP, 3) & "', " & _
                                " " & RsPromedio!PrecioPromedio & ", '" & sPorcFR & "' )"
                        CnAlmacen.Execute Isql
                        'Fin 16/03/16
                        
                    End If
                End If
            End If
        
            RsDtemporal.MoveNext
        Next i
        RsDtemporal.MoveLast
    Else
        sFueraFluct = ""
        RsDocumentoD.MoveFirst
        For i = 0 To RsDocumentoD.RecordCount - 1
        
'            lConsoli = Calcular("SELECT isnull(lPPConsolidadoSA,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDocumentoD!tCodigoProducto & "')", CnAlmacen)
'            If cboMoneda.BoundText = "01" Then      'Moneda Nacional
'                If lPPConsolidado Then
'                    If lConsoli Then
'                        Isql = "SELECT a.tCodigoProducto, a.nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                               & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + a.nCantidad * a.nPrecio) " _
'                               & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                               & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact,b.tDetallado Producto " _
'                               & " FROM dbo.DDOCUMENTO AS a INNER JOIN dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto left JOIN " _
'                               & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                               & " WHERE (b.tCodigoProducto = '" & RsDocumentoD!tCodigoProducto & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'                    Else
'                        Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                               & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * a.nprecio))" _
'                               & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact,b.tDetallado Producto From ddocumento a," _
'                               & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDocumentoD!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                    End If
'                Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * a.nprecio))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDocumentoD!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                End If
'            Else                                    'Moneda Extranjera
'                If lPPConsolidado Then
'                    If lConsoli Then
'                        Isql = "SELECT a.tCodigoProducto, a.nPrecio*c.ncambio as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                               & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + a.nCantidad * (a.nPrecio * c.nCambio)) " _
'                               & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                               & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact,b.tDetallado Producto " _
'                               & " FROM dbo.DDOCUMENTO AS a INNER JOIN " _
'                               & " dbo.MDOCUMENTO AS c ON a.tCorrelativo = c.tCorrelativo INNER JOIN " _
'                               & " dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto LEFT JOIN " _
'                               & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                               & " WHERE (b.tCodigoProducto = '" & RsDocumentoD!tCodigoProducto & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'
'                    Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio*c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (a.nPrecio * c.nCambio)))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDocumentoD!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'
'                    End If
'                Else
'                    Isql = "Select a.tcodigoproducto, a.nprecio*c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                           & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (a.nPrecio * c.nCambio)))" _
'                           & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact,b.tDetallado Producto From ddocumento a," _
'                           & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & RsDocumentoD!tCodigoProducto & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                End If
'            End If
        
            Isql = "exec usp_Almacen_CalculaPrecioPromedio '" & LblCorrelativo & "', '" & RsDocumentoD!tCodigoProducto & "', ''"
            
            Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)

            If RsPromedio.RecordCount > 0 Then
                RsPromedio.MoveFirst
                nPorcFN = Calcular("SELECT isnull(nPorcFlucNeg,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDocumentoD!tCodigoProducto & "')", CnAlmacen)

                nPorcFP = Calcular("SELECT isnull(nPorcFlucPos,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & RsDocumentoD!tCodigoProducto & "')", CnAlmacen)

                'RsPromedio!nPrecioPromedio
                If nPorcFN <> 0 Or nPorcFP <> 0 Then
                    nValPorcFN = RsPromedio!nPrecioPromedio - (nPorcFN * RsPromedio!nPrecioPromedio / 100)
                    nValPorcFP = RsPromedio!nPrecioPromedio + (nPorcFP * RsPromedio!nPrecioPromedio / 100)
                    If Not (nValPorcFN <= RsPromedio!PrecioPromedio And RsPromedio!PrecioPromedio <= nValPorcFP) And (nValPorcFN > 0 And nValPorcFP > 0) Then
                        sFueraFluct = sFueraFluct & vbCr & RsPromedio!tCodigoProducto & " - " & RsPromedio!Producto
                                                
                        'Jesus 16/03/16
                        nPorcFR = ((RsPromedio!PrecioPromedio - RsPromedio!nPrecioPromedio) * 100 / RsPromedio!nPrecioPromedio)
                        If nPorcFR < 0 Then
                            sPorcFR = "(- " & Round(Abs(nPorcFR), 2) & "%)"
                        Else
                            sPorcFR = "(+ " & Round(Abs(nPorcFR), 2) & "%)"
                        End If
                        
                        Isql = "insert into " & sTempoInsumos & " (tCodigo,tArticulo,nPrecioPromedio,nFlucNeg,nFlucPos,nPrecioPromedioNuevo,nFlugRes) values " & _
                                "('" & RsPromedio!tCodigoProducto & "','" & RsPromedio!Producto & "', " & RsPromedio!nPrecioPromedio & ", " & _
                                " '(- " & nPorcFN & "%) : " & Round(nValPorcFN, 3) & "', '(+ " & nPorcFP & "%) : " & Round(nValPorcFP, 3) & "', " & _
                                " " & RsPromedio!PrecioPromedio & ", '" & sPorcFR & "' )"
                        CnAlmacen.Execute Isql
                        'Fin 16/03/16
                        
                    End If
                End If
            End If
            
            RsDocumentoD.MoveNext
        Next i
        RsDocumentoD.MoveLast
    End If

    ValidaPrecioPromedio = sFueraFluct

End Function

Sub Consulta_Kardex(tipo As String)
    Dim lConsoli As Boolean
    'Precio Promedio ----------------------------------------------------------------------------
    sPromedio = 0
    
    If CboMoneda.BoundText <> "01" Then
        xPUni = xPUni * CDbl(LblTipoCambio)
    End If
    
    Isql = "exec usp_Almacen_CalculaPrecioPromedio '" & LblCorrelativo & "', '" & sCodi & "', '" & tipo & "'"
    Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsPromedio.EOF Then
        RsPromedio.MoveLast
        sUltPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
    Else
        sUltPromedio = 0
    End If
    sPromedio = sUltPromedio
                
'    lConsoli = Calcular("SELECT isnull(lPPConsolidadoSA,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & sCodi & "')", CnAlmacen)
'    If cboMoneda.BoundText = "01" Then      'Moneda Nacional
'        If lPPConsolidado Then
'            If lConsoli Then
'                Isql = "SELECT a.tCodigoProducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                       & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + (a.nCantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))) " _
'                       & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                       & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact " _
'                       & " FROM dbo.DDOCUMENTO AS a INNER JOIN dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto left JOIN " _
'                       & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                       & " WHERE (b.tCodigoProducto = '" & sCodi & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'            Else
'                Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                       & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0))))" _
'                       & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact From ddocumento a," _
'                       & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'            End If
'        Else
'            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                   & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0))))" _
'                   & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (a.ncantidad + b.nstockactual) As stockact From ddocumento a," _
'                   & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'        End If
'    Else                                    'Moneda Extranjera
'       'DETP 15/03/2006 No se debe multiplicar el precio promedio del tproducto si este ya se encuentra en soles
'        If lPPConsolidado Then
'            If lConsoli Then
'                Isql = "SELECT a.tCodigoProducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                       & " (case when (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad)=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) + a.nCantidad * (c.nCambio * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))) " _
'                       & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.ncantidad) end) AS PrecioPromedio," _
'                       & " ((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))+a.nCantidad) AS stockact " _
'                       & " FROM dbo.DDOCUMENTO AS a INNER JOIN " _
'                       & " dbo.MDOCUMENTO AS c ON a.tCorrelativo = c.tCorrelativo INNER JOIN " _
'                       & " dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto LEFT JOIN " _
'                       & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                       & " WHERE (b.tCodigoProducto = '" & sCodi & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'
'            Else
'            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                   & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (c.nCambio * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))))" _
'                   & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact From ddocumento a," _
'                   & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'
'            End If
'        Else
'            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                   & "(case when (abs(b.nstockactual)+a.ncantidad)=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)+(a.ncantidad * (c.nCambio * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))))" _
'                   & " /(abs(b.nstockactual)+a.ncantidad)) end) As PrecioPromedio, (b.nstockactual+a.ncantidad) As stockact From ddocumento a," _
'                   & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'        End If
'        xPUni = xPUni * CDbl(LblTipoCambio)
'    End If
'
'    Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
'    If Not RsPromedio.EOF Then
'        RsPromedio.MoveLast
'        sPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
'    End If
'
'    If Tipo = "D" Then
'
'                lConsoli = Calcular("SELECT isnull(lPPConsolidadoSA,0) as codigo from tsubfamilia where tCodigoSubFamilia=(SELECT tCodigoSubFamilia FROM TPRODUCTO WHERE tcodigoproducto='" & sCodi & "')", CnAlmacen)
'                If cboMoneda.BoundText = "01" Then      'Moneda Nacional
'                    If lPPConsolidado Then
'                        If lConsoli Then
'                            Isql = "SELECT a.tCodigoProducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                                   & " (case when CAST((abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) - (a.nCantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))) " _
'                                   & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.ncantidad) end) AS PrecioPromedio," _
'                                   & " CAST(((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.nCantidad) AS DECIMAL(18,3)) AS stockact " _
'                                   & " FROM dbo.DDOCUMENTO AS a INNER JOIN dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto LEFT JOIN " _
'                                   & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                                   & " WHERE (b.tCodigoProducto = '" & sCodi & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'                        Else
'                            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                                   & "(case when CAST((abs(b.nstockactual)-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)-(a.nCantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0))))" _
'                                   & " /(abs(b.nstockactual)-a.ncantidad)) end) As PrecioPromedio, CAST((b.nstockactual-a.ncantidad) AS DECIMAL(18,3)) As stockact From ddocumento a," _
'                                   & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                        End If
'                    Else
'                            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                                   & "(case when CAST((abs(b.nstockactual)-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)-(a.nCantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0))))" _
'                                   & " /(abs(b.nstockactual)-a.ncantidad)) end) As PrecioPromedio, CAST((b.nstockactual-a.ncantidad) AS DECIMAL(18,3)) As stockact From ddocumento a," _
'                                   & " tproducto b Where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                    End If
'                Else                                    'Moneda Extranjera
'                   'DETP 15/03/2006 No se debe multiplicar el precio promedio del tproducto si este ya se encuentra en soles
'                    If lPPConsolidado Then
'                        If lConsoli Then
'
'                            Isql = "SELECT a.tCodigoProducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.nCantidad, b.nStockActual, b.nPrecioPromedio, " _
'                                   & " (case when CAST((abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else (ABS((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))* b.nPrecioPromedio) - (a.nCantidad * (c.nCambio * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0))))) " _
'                                   & " / (abs(b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.ncantidad) end) AS PrecioPromedio," _
'                                   & " CAST(((b.nStockActual + isnull(dbo.vStockSubAlmacenes.nStockActualSA,0))-a.nCantidad) AS DECIMAL(18,3)) AS stockact " _
'                                   & " FROM dbo.DDOCUMENTO AS a INNER JOIN " _
'                                   & " dbo.MDOCUMENTO AS c ON a.tCorrelativo = c.tCorrelativo INNER JOIN " _
'                                   & " dbo.TPRODUCTO AS b ON a.tCodigoProducto = b.tCodigoProducto LEFT JOIN " _
'                                   & " dbo.vStockSubAlmacenes ON b.tCodigoProducto = dbo.vStockSubAlmacenes.tCodigoProducto " _
'                                   & " WHERE (b.tCodigoProducto = '" & sCodi & "') AND (a.tCorrelativo = '" & LblCorrelativo & "') "
'                        Else
'
'                            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                                   & "(case when CAST((abs(b.nstockactual)-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)-(a.ncantidad * (c.nCambio * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)))))" _
'                                   & " /(abs(b.nstockactual)-a.ncantidad)) end) As PrecioPromedio, CAST((b.nstockactual-a.ncantidad) AS DECIMAL(18,3)) As stockact From ddocumento a," _
'                                   & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'
'                        End If
'                    Else
'                            Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * Isnull(a.nPorcentajeLeyAD,0)) * c.ncambio as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, " _
'                                   & "(case when CAST((abs(b.nstockactual)-a.ncantidad) AS DECIMAL(18,3))=0 then b.nPrecioPromedio else ((abs(b.nstockactual * b.nPrecioPromedio)-(a.ncantidad * (c.nCambio * (a.nPrecio - (a.nPrecio * Isnull(a.nPorcentajeLeyAD,0))))))" _
'                                   & " /(abs(b.nstockactual)-a.ncantidad)) end) As PrecioPromedio, CAST((b.nstockactual-a.ncantidad) AS DECIMAL(18,3)) As stockact From ddocumento a," _
'                                   & " tproducto b, mdocumento c Where c.tCorrelativo=a.tCorrelativo And a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'
'                    End If
'                    xPUni = xPUni * CDbl(LblTipoCambio)
'                End If
'                Debug.Print Isql
'                Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
'                If Not RsPromedio.EOF Then
'                    RsPromedio.MoveLast
'                    sUltPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
'                Else
'                    sUltPromedio = 0
'                End If
'
'          sPromedio = sUltPromedio
'    End If

    'Linea Agregada por Luis
    Isql = "sp_TraeDatosConsultaKardex '" & sCodi & "', '" & CboArea.BoundText & "'"
    Set RsDatos = Lib.OpenRecordset(Isql, CnAlmacen)
    With RsDatos
        If .RecordCount > 0 Then
            sUltimo = !stkProducto
            sAnteriorKD = !stkProducto
            sPCosto = !PrecioCosto
            sAnteriorSKD = !StkSubStock
            xCodigo = !codProducto
            xArea = !CodArea
            CorrKD = !CorrmKardex
            CorrSKD = !CorrmSubKardex
            CorrSTK = !CorrSubStock
        Else
            sUltimo = 0
            sAnteriorKD = 0
            sPCosto = 0
            sAnteriorSKD = 0
            xCodigo = ""
            xArea = ""
            CorrKD = 0
            CorrSKD = 0
            CorrSTK = 0
        End If
    End With
    'Fin Linea Agregada por Luis

    sSub_Stock = sAnteriorSKD

    'Buscar Stock en Tabla Requerimiento ----------------------------------------------------------------------------
    Isql = "Select tPrt.nStockActual,tPrt.tCodigoProducto" _
           & " From DRequerimiento DReq Right Outer Join " _
           & " tProducto tPrt On DReq.tCodigoProducto = tPrt.tCodigoProducto" _
           & " Where tPrt.tCodigoProducto='" & sCodi & "'"
    Set RsUltRequerimiento = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsUltRequerimiento.EOF Then
        RsUltRequerimiento.MoveLast
        sUltRequerimiento = IIf(IsNull(RsUltRequerimiento!nStockActual), 0, RsUltRequerimiento!nStockActual)
    End If
    'Busco Anular Correlativo en Tabla Requerimiento ----------------------------------------------------------------------------
    Isql = "SELECT MReq.tRequerimiento FROM MDocumento MDoc," _
           & " MRequerimiento MReq Where MDoc.tDocumento = MReq.tDocumento" _
           & " AND MDoc.tCorrelativo='" & LblCorrelativo & "'"
    Set RsAnula = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsAnula.EOF Then
        RsAnula.MoveLast
        lsAnula = IIf(IsNull(RsAnula!tRequerimiento), "", RsAnula!tRequerimiento)
    End If

    CorrKD = IIf(CorrKD = 0, 1, CorrKD)
    'Busco Maximo Correlativo en Tabla MSubKardex ----------------------------------------------------------------------------
    CorrSKD = IIf(CorrSKD = 0, 1, CorrSKD)
    'Busco Maximo Correlativo en Tabla tSubStock ----------------------------------------------------------------------------
    CorrSTK = IIf(CorrSTK = 0, 1, CorrSTK)
     If lGeneraNS Then   'nuevo
        'Modificado por Luis
        xAnho = Year(fFechaProcesa)
        xMes = Mid("00", 1, 2 - Len(Trim(str(Month(fFechaProcesa))))) & Trim(str(Month(fFechaProcesa)))
        'Fin Modificacion
        xPeriodo = Mid(xAnho, 3) & xMes
        Set Rs = Lib.OpenRecordset("sp_TraeMaximoRequerimiento '" & xPeriodo & "'", CnAlmacen) 'Error... reiniciaba el correlativo
        If Rs.RecordCount > 0 Then
            nCorrela = IIf(IsNull(Rs!codigo), 0, Rs!codigo)
        Else
            nCorrela = 0
        End If
        If IsNull(nCorrela) Or nCorrela = "" Or nCorrela = "0" Then
            sBRequerimiento = xPeriodo & "000001"
        Else
            sBRequerimiento = xPeriodo & Lib.Correlativo(nCorrela, 6)
        End If
    End If
End Sub

Private Sub cmdPago_Click()
    Set mmPago = New ClsDocumento
    If Calcular("select lCtaCte as Codigo from vTipoDocumento where Codigo='" & CboTipoDocumento.BoundText & "'", CnAlmacen) = False Then
        MsgBox "Este documento no permite pago", vbInformation, sMensaje
        Exit Sub
    End If
        Select Case pais
            Case "000", "004", "002", "003", "005"
'                    If MsgBox("El Documento se pagara en:" & Chr(13) & "Moneda: " & CboMoneda.Text & Chr(13) & "  Monto: " & Format(frmDocumento.RsCabecera!TOTAL + frmDocumento.RsCabecera!nPercepcion, "##,###,##0.00") & Chr(13) & " ¿Desea continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
'                        Exit Sub
'                    End If
            Case "001"
'                    If MsgBox("El Documento se pagara en:" & Chr(13) & "Moneda: " & CboMoneda.Text & Chr(13) & "  Monto: " & Format(frmDocumento.RsCabecera!CabSaldo, "##,###,##0.00") & Chr(13) & " ¿Desea continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
'                        Exit Sub
'                    End If
        End Select

'    If MsgBox("El Documento se pagara en:" & Chr(13) & "Moneda: " & CboMoneda.Text & Chr(13) & "  Monto: " & Format(FrmDocumento.RsCabecera!total + FrmDocumento.RsCabecera!nPercepcion, "##,###,##0.00") & Chr(13) & " ¿Desea continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
'        Exit Sub
'    End If
'
    If nTC = 0 Then MsgBox "No se puede realizar el pago, ya que el tipo de cambio es ''0''.", vbInformation + vbOKOnly, "Sistemas": Exit Sub
    
    On Error GoTo ErrorPago
    Screen.MousePointer = vbHourglass
    CnAlmacen.BeginTrans
    
    If lAuditoria Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "MDOCUMENTO", "INGRESO DE DOCUMENTOS", "02", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), LblCorrelativo, "", "tCorrelativo", "Correlativo", LblCorrelativo, "tEstadoDocumento", "Estado de Documento", "03", "fCancelado", "Fecha de Cancelación", FechaServidor())

        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    
    mmPago.InsmPago TxtProveedor.Text, "", sDocumento, "", LblTotal.Caption, CboMoneda.BoundText, nTC, tPagoRapidoTipoPago, sUsuario, "", "", "", "", "", "NULL", 1, LblCorrelativo, "A"
    
   
    Set mmPago = Nothing
    CnAlmacen.CommitTrans
'    frmDocumento.RsCabecera.Requery
'    If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.MoveFirst
'    frmDocumento.RsCabecera.Find "tCorrelativo='" & LblCorrelativo & "'"      'Ricky
    Screen.MousePointer = vbDefault
    'MsgBox "Documento Pagado", vbInformation, sMensaje
'    Asignar
    Exit Sub
ErrorPago:
    CnAlmacen.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrio un error al tratar de pagar", vbInformation, sMensaje
End Sub

Private Sub CmdProcesar_Click()
    Dim nLen As Integer
    Dim fProce As Date
    Dim sProdFueraRangoFluct As String
    Dim FechaContable As Date
    Dim oDiaContable As New clsDiaContable
    
    FechaContable = oDiaContable.obtieneDiaContable
    
    'If FunAccesoBoton(sCodProyecto, sUsuario, "00091", "05") = False Then Exit Sub
    
    If grdGrilla.VisibleRows = 0 Then
        MsgBox "No existe datos para " & CmdProcesar.Caption & "...", vbInformation, sMensaje
        Exit Sub
    End If

    ValidaCierre
    If sCerrado = "Verdadero" Then Exit Sub
    
    ValidaPlazo
    'If Not swPasar Then Exit Sub
      
    sPromedio = 0
    sUltPromedio = 0
    sUltimo = 0
    sAnteriorKD = 0
    sAnteriorSKD = 0
    sPCosto = 0
    sSub_Stock = 0
    sUltRequerimiento = ""
    lsAnula = ""
    CorrKD = 0
    CorrSKD = 0
    CorrSTK = 0
    sBRequerimiento = ""
   
'    If CmdProcesar.Caption = "Desprocesar" Then
'        Set mdDocumento = New ClsDocumento
'        If mdDocumento.VerificaModificacion(LblCorrelativo.Caption) > 0 Then
'            MsgBox "No es posible desprocesar el documento. " & vbCrLf _
'                       & "Se a modificado el detalle del documento...", vbCritical, sMensaje
'            wEnter = False
'            Exit Sub
'        End If
'
'        If CDbl(LblTotal) + CDbl(lblTotLey1) > CDbl(LblSaldo) Then
'            If Calcular("select count(tDocumento) as Codigo from mPago where tCodigoProveedor='" & Trim(TxtProveedor.Text) & "' and tDocumento='" & Trim(LblPDoc.Caption) & Trim(TxtSerie.Text) & Trim(TxtCorrelativo.Text) & "'", CnAlmacen) > 0 Then
'                MsgBox "No es posible desprocesar el documento. Hay pagos rea-" & vbCrLf _
'                       & "lizados o Notas de Crédito aplicadas... (ver Historial)", vbCritical, sMensaje
'                wEnter = False
'                Exit Sub
'            End If
'        End If
'        Isql = "select count(tDocumento) as Codigo from mPago where tCodigoProveedor='" & Trim(TxtProveedor.Text) & "' and tDocumento='" & Trim(LblPDoc.Caption) & Trim(TxtSerie.Text) & Trim(TxtCorrelativo.Text) & "'"
'
'        If sReqDoc <> "" Then
'            Isql = "select tCodigoPorcionamiento from mporcionamiento where lAutomatico=1 and tEstado<>'04' and tObservacion='" & sReqDoc & "'"
'            Set RsSearch = Lib.OpenRecordset(Isql, CnAlmacen)
'            If Not RsSearch.EOF Then MsgBox "No es posible desprocesar este documento." & vbCrLf & "Para hacerlo, anule el porcionamiento " & RsSearch!tCodigoPorcionamiento & "...", vbCritical, sMensaje: Exit Sub
'        End If
'        '~~~~~ Verifica si el documento a desprocesar es una guía canjeada ~~~~~
'        If CboTipoDocumento.BoundText = "46" And LblPDocRef <> "" And TxtSerieRef <> "" And TxtCorrelativoRef <> "" Then
'            MsgBox "No es posible desprocesar. Hay un documento asociado" & vbCrLf & _
'                   "a esta guía. Revierta el canje de ésta seleccionando" & vbCrLf & _
'                   "el documento " & Trim(LblPDocRef) & Trim(TxtSerieRef) & Trim(TxtCorrelativoRef) & "...", vbCritical, sMensaje
'            wEnter = False
'            Exit Sub
'        End If
'
'        '~~~~~ Verifica si el documento a desprocesar está canjeado con guía(s) ~~~~~
'        Isql = "SELECT TDOCUMENTO FROM MDOCUMENTO WHERE TTIPODOCUMENTO = '46' AND TDOCUMENTOREFERENCIA = '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' AND TCODIGOPROVEEDOR = '" & TxtProveedor & "'"
'        Set RsTemp = Lib.OpenRecordset(Isql, CnAlmacen)
'        If RsTemp.RecordCount = 0 Then
'            Set RsTemp = Nothing
'        Else
'            Set RsTemp = Nothing
'            MsgBox "No es posible desprocesar. Hay una o más guías" & vbCrLf & _
'                   "asociadas a este documento. Revierta el canje...", vbCritical, sMensaje
'            wEnter = False
'            Exit Sub
'        End If
'
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''        If Contraseña("05") = False Then
''            MsgBox "Clave no permitida", vbExclamation, sMensaje
''            Exit Sub
''        End If
'
'        Desprocesar
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'If lLibroCompra And dtcLibroCompra = "" Then MsgBox "Ingrese el Libro de Compra Antes de Procesar...", vbExclamation, sMensaje: dtcLibroCompra.SetFocus: Exit Sub

    sUsuarioE = ""
'    If lAdvertenciaDifFechas Then
'        fProce = CDate(Format(FechaServidor, "dd/mm/yyyy"))
'        If DtpIngreso.value <> fProce Then
'           frmContrasenaUsuario.tOpc = "2"
'           frmContrasenaUsuario.Show vbModal
'
'           If frmContrasenaUsuario.wEnter = False Then
'               Exit Sub
'           Else
'               cHistorial.FunInsertatHistorial "001", "009", "", LblCorrelativo, "", "", "", "", sUsuario, sUsuarioE, "Las Fechas de Ingreso y Proceso del documento " & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & " fueron validadas."
'           End If
'        End If
'    End If
    sUsuarioE = ""
    If CDbl(lblGranTotal) <> CDbl(LblSaldo) Then
        MsgBox "No es posible procesar..." & vbCrLf & "Antes, grabe el documento", vbCritical, sMensaje
        cmdOpcion(4).SetFocus
        Exit Sub
    End If
    
    'If MsgBox("¿Seguro de procesar el documento " & LblPDoc.Caption & TxtSerie & TxtCorrelativo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then Exit Sub
    If MensajeCabecera Then
        Exit Sub
    End If
    
    If ValidaDuplicidad Then
        MsgBox "El Documento ha cambiado de estado, debe verificar otras estaciones", vbExclamation, sMensaje
        Exit Sub
    End If
    
    'Verifica precios minimos y maximos
    'Se cambio para el procesar
'    sProdFueraRangoFluct = ValidaPrecioPromedio()
'    If sProdFueraRangoFluct <> "" Then
'
'        'Jesus 16/03/16
'        'MsgBox "El precio promedio resultante de los artículos: " & vbCr & sProdFueraRangoFluct & vbCr & vbCr & "está FUERA DEL RANGO DE FLUCTUACION DEL PRECIO PROMEDIO.", vbExclamation + vbOKOnly, sMensaje
'        frmVerificaRangoFluctuacionPP.tFormulario = "D"
'        frmVerificaRangoFluctuacionPP.Show vbModal
'        'Fin 16/03/16
'
'        If Contraseña("10") = False Then
'            MsgBox "Clave no permitida", vbExclamation, sMensaje
'            Exit Sub
'        End If
'    End If
    
    Isql = "sp_VerificaPreciosMinimosMaximos '" & Trim(LblCorrelativo.Caption) & "'"
    Set RsTemp = Lib.OpenRecordset(Isql, CnAlmacen)
    nCount = 0
    If RsTemp.RecordCount > 0 Then
        nCount = RsTemp.RecordCount
        RsTemp.MoveFirst
        sProValida = ""
        Do While Not RsTemp.EOF
            sProValida = sProValida & IIf(sProValida <> "", Chr(13), "") & RsTemp!tResumido & " (Precio Min.: " & Format(RsTemp!nPrecioMinimo, "##,##0.000") & " Precio Art: " & Format(RsTemp!nPrecio, "##,#0.000") & " Precio Max.: " & Format(RsTemp!nPrecioMaximo, "##,##0.000") & "), "
            RsTemp.MoveNext
        Loop
        Set RsTemp = Nothing
        nLen = Len(Trim(sProValida)) - 1
        sProValida = Left(Trim(sProValida), nLen)
        If nCount = 1 Then
            MsgBox "El articulo: " & Chr(13) & _
                   sProValida & Chr(13) & _
                   "no se encuentra dentro del Rango de su PRECIO DE COSTO minimo y/o maximo", vbInformation, sMensaje
        Else
            MsgBox "Los articulos: " & Chr(13) & _
                   sProValida & Chr(13) & _
                   "no se encuentran dentro del Rango de sus PRECIO DE COSTO minimos y/o maximos", vbInformation, sMensaje
        End If
        
        Screen.MousePointer = vbDefault
'        If Contraseña("07") = False Then
'            MsgBox "Clave no permitida", vbExclamation, sMensaje
'            Exit Sub
'
'        End If
        
'        lObserva = False
'        frmObservacion.Show vbModal
'        If lObserva = False Then
'            MsgBox "La contraseña no es válida...", vbCritical, sMensaje
'            Exit Sub
'        End If
        Set mDocObservado = New ClsDocumento
        'REVISAR
        'mDocObservado.DocumentoObservacion sDocumento, frmDocumento.RsCabecera!tcodigoproveedor, Format(FechaServidor, "YYYY/mm/dd HH:mm:ss"), frmDocumento.RsCabecera!tCorrelativo, True, "I"
        
        Set mDocObservado = Nothing
    End If
    
    'Jesus 15-11-04
    'Se agrego validacion de estado despues de la validacion de fluctuacion y precios maximos y minimos
    If ValidaDuplicidad Then
        MsgBox "El Documento ha cambiado de estado, debe verificar otras estaciones", vbExclamation, sMensaje
        Exit Sub
    End If
    'Fin Cambio

    Set RsTemp = Nothing
    If Calcular("Select lFecha as codigo from tParametro", CnAlmacen) = True Then
        fFechaProcesa = Format(Calcular("select getdate() as Codigo", CnAlmacen), "YYYY/mm/dd HH:mm:ss")
    Else
        fFechaProcesa = Format(Calcular("select getdate() as Codigo", CnAlmacen), "YYYY/mm/dd HH:mm:ss")
    End If
    
    'FechaTransaccion = Format(FrmDocumento.RsCabecera!FecIngreso, "YYYY/mm/dd") & " " & Format(FechaServidor(), "HH:mm:ss")
    FechaTransaccion = Format(DtpIngreso.value, "YYYY/mm/dd HH:mm") ' & " 00:00:00"
'    If lKardexFechaIngreso Then
'        FechaTransaccion = Format(ComparaFecha(DtpIngreso.value), "YYYY/mm/dd HH:mm") '& " 00:00:00"
'        If lPermiteFecha = False Then Exit Sub
'        If (ValidaCierrePosterior(FechaTransaccion, CboArea.BoundText, "", CboTipoDocumento.BoundText)) > 0 Then
'            MsgBox "Existe un cierre de inventario posterior a la fecha de ingreso del documento. ", vbExclamation, sMensaje
'            Exit Sub
'        End If
'    End If
    
    
    Screen.MousePointer = vbHourglass
    Set mdDocumento = New ClsDocumento
    
    If lAuditoria Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "MDOCUMENTO", "INGRESO DE DOCUMENTOS", "02", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), LblCorrelativo, "", "tCorrelativo", "Correlativo", LblCorrelativo, "tEstadoDocumento", "Estado de Documento", "02", "tProcesa", "Usuario que Procesó", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), "fProcesa", "Fecha de Proceso", FechaServidor())
    
        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    'Jesus 15/11/05
    On Error GoTo ErrorProceso
    CnAlmacen.BeginTrans
    'Fin
    
    mdDocumento.CambiaEstadoDocumento LblCorrelativo, sUsuario, "02", Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss"), Format(DtpEmision.value, "YYYY/mm/dd"), 2, Format(FechaContable, "yyyy/mm/dd")
    Set mdDocumento = Nothing
    
    swEntrar = False
    'Asignar
    swEntrar = True
    
    sCodi = ""
    
   '  Solamente se genera el movimiento de kardex, stocks y precios si el tipo de documento lo requiere...
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If CboArea.BoundText = "000" Then        'Almacen Central
        If Sw Then
            MsgBox "Ocurrio un errór en el proceso. ", vbCritical, sMensaje
            Exit Sub
        Else  'if sw
            RsDocumentoD.MoveFirst
            Do While Not RsDocumentoD.EOF
                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                sValor = IIf(Not IsNull(RsDocumentoD!nNetoAD), RsDocumentoD!nNetoAD, "")
                Select Case pais
                    Case "000", "002", "003", "004", "005"
                        xPUni = CDbl(IIf(Not IsNull(RsDocumentoD!nPrecioAD), RsDocumentoD!nPrecioAD, "0"))
                    Case "001"
                        xPUni = CDbl(IIf(Not IsNull(RsDocumentoD!nPrecioAD), RsDocumentoD!nPrecioAD, "0"))
                End Select
                sAreaO = IIf(Not IsNull(RsDocumentoD!tAreaOrigen), RsDocumentoD!tAreaOrigen, "")
              
                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                    Consulta_Kardex "P"
                    CorrKD = CorrKD + 1
                    
                    Isql = "Update tProducto Set " _
                           & "nPrecioCosto =" & xPUni & "," _
                           & "nPrecioPromedio =" & sPromedio & "," _
                           & "nStockActual=" & CDbl(sUltimo) + CDbl(sCant) & "," _
                           & "lreplica=1  " _
                           & "Where tCodigoProducto='" & sCodi & "'"
                    CnAlmacen.Execute Isql
                    
                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
                
                   'Ingreso del Artículo...
                    Isql = "Insert Into mKardex (" _
                           & " tCodigoProducto,nCorrelativo,tTipoDocumento,tCodigoArea,tDocumento," _
                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
                           & " values ('" & sCodi & "', " _
                           & " '" & CorrKD & "', " _
                           & " '" & CboTipoDocumento.BoundText & "', " _
                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & xPUni * CDbl(sCant) & ", " _
                           & "" & 0 & ", " _
                           & "" & CDbl(sUltimo) + CDbl(sCant) & ", " _
                           & "" & CDbl(sUltimo) & ", " _
                           & "'01'," & " '" & IIf(sPromedio = 0, sCosto, sPromedio) & "', '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
                           & " '" & sUsuario & "', " _
                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
                    CnAlmacen.Execute Isql
                End If
                RsDocumentoD.MoveNext
            Loop
        End If
    Else 'Sub-Almacenes
        sPrf = "RQ"
        sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
        If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
            GeneraEncargado
            lGeneraNS = True
            Consulta_Kardex "P"
            Isql = "Insert Into mRequerimiento (" _
                   & "tRequerimiento,tCodigoArea,tResponsable," _
                   & "tEstadoRequerimiento,tUsuario,fRegistro," _
                   & "tProcesa,fProcesa,tDocumento,tCorrelativo,lReplica,fDiaContable,lPedido) " _
                   & "Values ('" & Trim(sBRequerimiento) & "', " _
                   & "'" & CboArea.BoundText & "', " & " '" & Trim(sEncargado) & "','02'," _
                   & "'" & sUsuario & "', '" & Format(IIf(lKardexFechaIngreso, FechaTransaccion, fFechaProcesa), "YYYY/mm/dd HH:mm:ss") & "', " _
                   & "'" & sUsuario & "', '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
                   & "'" & Trim(LblPDoc & TxtSerie & TxtCorrelativo) & "'," _
                   & "'" & LblCorrelativo & "' , 1,convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),0)"
            CnAlmacen.Execute Isql

            '~~~ Graba el Requerimiento automático en la cabecera de Documentos ~~~
            Isql = "Update mDocumento SET " _
                   & " tRequerimiento ='" & Trim(sBRequerimiento) & "'" _
                   & " ,lreplica=1 WHERE tCorrelativo='" & LblCorrelativo & "'"
            CnAlmacen.Execute Isql
        End If

        If Sw Then
            MsgBox "Ocurrio un errór en el proceso. ", vbCritical, sMensaje
            Exit Sub
        Else
            RsDocumentoD.MoveFirst
            Do While Not RsDocumentoD.EOF
                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                sValor = IIf(Not IsNull(RsDocumentoD!nNetoAD), RsDocumentoD!nNetoAD, "")
                nDItem = RsDocumentoD!nItem
                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)

                    Isql = "Select a.tcodigoproducto, ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * 1) as nPrecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, (((b.nstockactual * b.nPrecioPromedio)+ " _
                           & " (a.ncantidad * ((a.nPrecio / (1 + Isnull(a.nPorcentajeLeyAD,0))) * 1)))/ (CASE a.ncantidad + b.nstockactual WHEN 0 THEN 1 ELSE a.ncantidad + b.nstockactual END)) as PrecioPromedio, (a.ncantidad + b.nstockactual) as stockact,b.nPrecioCosto from ddocumento a," _
                           & " tproducto b where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
                    Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
                    If Not RsPromedio.EOF Then
                        RsPromedio.MoveLast
                        sPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
                    End If

                    If sUltRequerimiento = "" Then sUltRequerimiento = 0
                    Isql = "Insert Into dRequerimiento (" _
                           & "tRequerimiento,tCodigoProducto," _
                           & " nCantidadPedida,nAtendido,nPendiente,nStockActual,nPrecioPromedio,nItem)" _
                           & " Values ('" & Trim(sBRequerimiento) & "', " _
                           & " '" & sCodi & "', " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & 0 & ", " _
                           & "" & CDbl(sUltRequerimiento) + CDbl(sCant) & ", " _
                           & "" & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & "," & nDItem & ")"
                    CnAlmacen.Execute Isql

'                    If Calcular("select lPorcionablePL as codigo from tParametro", CnAlmacen) = True Then
'                        Set clsPorciona = New clsPorcionamiento
'                        Isql = "SELECT tArticuloPorcionado.tcodigoarticuloRes,dbo.tArticuloPorcionado.tcodigoPorcionable,dbo.tArticuloPorcionado.tcodigoPorcionado, dbo.tArticuloPorcionado.nPorcentajeMerma, dbo.tArticuloPorcionado.nEquivalencia, dbo.TPRODUCTO.tDetallado FROM dbo.tArticuloPorcionado INNER JOIN dbo.TPRODUCTO ON dbo.tArticuloPorcionado.tCodigoPorcionado = dbo.TPRODUCTO.tCodigoProducto WHERE  (dbo.tArticuloPorcionado.lPorcAutomatico = 1) and tcodigoporcionable='" & sCodi & "'"
'                        Set rsPlantilla = Lib.OpenRecordset(Isql, CnAlmacen)
'                        If rsPlantilla.RecordCount > 0 Then
'                            rsPlantilla.MoveFirst
'                            sCodPorciona = clsPorciona.InsUpdPorciona("", sCodi, CboArea.BoundText, sUsuario, CDbl(sCant), CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), "01", "NULL", "NULL", Trim(sBRequerimiento), "I", True, Format(IIf(lKardexFechaIngreso, FechaTransaccion, fFechaProcesa), "YYYY/mm/dd HH:mm:ss"))
'                            Isql = "Insert into dPorcionamiento (tCodigoPorcionamiento, tCodigoProducto, tDetallado, tCantidadProducto, tEquivalencia, nPProm)  " & _
'                                   " values ('" & sCodPorciona & "','" & rsPlantilla!tCodigoPorcionado & "','" & rsPlantilla!tDetallado & "'," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & "," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & ", 0)"
'                            CnAlmacen.Execute Isql
'                        End If
'                        Set clsPorciona = Nothing
'                    End If
                End If
                RsDocumentoD.MoveNext
            Loop
        End If

        If Sw Then
            MsgBox "Ocurrio un errór en el proceso. ", vbCritical, sMensaje
            Exit Sub
        Else                      'Se está modificando...
            RsDocumentoD.MoveFirst
            Do While Not RsDocumentoD.EOF
                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                xPUni = IIf(Not IsNull(RsDocumentoD!nPrecioAD), RsDocumentoD!nPrecioAD, "")
                sAreaO = IIf(Not IsNull(RsDocumentoD!tAreaOrigen), RsDocumentoD!tAreaOrigen, "")
                sValor = IIf(Not IsNull(RsDocumentoD!nNetoAD), RsDocumentoD!nNetoAD, "")
                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)

                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                    lGeneraNS = False
                    Consulta_Kardex "P"
                    CorrKD = CorrKD + 1

                    'Ingreso
                    Isql = "Insert Into mKardex (" _
                           & "tCodigoProducto,nCorrelativo,tTipoDocumento,linterno,tCodigoArea,tDocumento," _
                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
                           & " Values ('" & sCodi & "', " _
                           & " '" & CorrKD & "','" & CboTipoDocumento.BoundText & "','1'," _
                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & CDbl(xPUni) * CDbl(sCant) & ", " _
                           & "" & 0 & ", " & " " & CDbl(sAnteriorKD) + CDbl(sCant) & ", " _
                           & "" & CDbl(sAnteriorKD) & ", " _
                           & "'01'," & " " & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & ", '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
                           & " '" & sUsuario & "', " _
                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
                    CnAlmacen.Execute Isql
                    'DETP se agrego el precio de promedio 10/08/2007
                    Isql = "update tProducto set nStockActual=isnull(nStockActual,0)+" & CDbl(sCant) & "" & IIf(CDbl(sPromedio) = 0, "", " , nPrecioPromedio =" & sPromedio & "") & ",lreplica=1  where tCodigoProducto='" & sCodi & "'"
                    CnAlmacen.Execute Isql
                    GeneraSalida     'Genera la salida del item antes de pasar al otro...
                End If

                RsDocumentoD.MoveNext
            Loop
        End If
    End If
      
   

   
    'cHistorial.FunInsertatHistorial "001", "007", "", LblCorrelativo, "", "", "", "", sUsuario, sUsuarioAutoriza, "Procesó el documento " & Trim(LblPDoc) & Trim(txtSerie) & Trim(txtCorrelativo) & ""
   
'    CmdProcesar.Caption = "Desprocesar"

    Set mdDocumento = New ClsDocumento
    mdDocumento.HistorialCompra LblCorrelativo, "I"
    Set mdDocumento = Nothing
    
    'Jesus 15/11/05
    CnAlmacen.CommitTrans
    CmdProcesar.Caption = "Desprocesar"
    'Fin

    'cHistorial.FunInsertatHistorial "001", "007", "", LblCorrelativo, "", "", "", "", sUsuario, sUsuarioAutoriza, "Procesó el documento " & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & ""


    Screen.MousePointer = vbDefault
    'MsgBox "Información procesada...", vbInformation, sMensaje
    
    'Dim iRow As Integer
'    frmDocumento.RsCabecera.Requery
'    If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.MoveFirst
'    frmDocumento.RsCabecera.Find "tCorrelativo='" & LblCorrelativo & "'"
'
    swEntrar = False
    'Asignar
    swEntrar = True
    
    ActivaAdicion False
    ControlEstado
    grdGrilla.MoveFirst
    sUsuarioAutoriza = ""
    Exit Sub

ErrorProceso:
    CnAlmacen.RollbackTrans
    MsgBox "No se logro Procesar Información. " & Error, vbCritical, sMensaje
    Screen.MousePointer = vbDefault
End Sub

Public Sub ProcesoAutomatico()
    Dim nLen As Integer
    Dim fProce As Date
    Dim sProdFueraRangoFluct As String
    Dim FechaContable As Date
    Dim oDiaContable As New clsDiaContable
    
    FechaContable = oDiaContable.obtieneDiaContable
    
    'Verifica precios minimos y maximos
    'Se cambio para el procesar
'    sProdFueraRangoFluct = ValidaPrecioPromedio()
'    If sProdFueraRangoFluct <> "" Then
'        MsgBox "El precio promedio resultante de los artículos: " & vbCr & sProdFueraRangoFluct & vbCr & vbCr & "está FUERA DEL RANGO DE FLUCTUACION DEL PRECIO PROMEDIO.", vbExclamation + vbOKOnly, sMensaje
'        If Contraseña("10") = False Then
'            MsgBox "Clave no permitida", vbExclamation, sMensaje
'            Exit Sub
'        End If
'    End If
    
    Isql = "sp_VerificaPreciosMinimosMaximos '" & Trim(LblCorrelativo.Caption) & "'"
    Set RsTemp = Lib.OpenRecordset(Isql, CnAlmacen)
    nCount = 0
    If RsTemp.RecordCount > 0 Then
        nCount = RsTemp.RecordCount
        RsTemp.MoveFirst
        sProValida = ""
        Do While Not RsTemp.EOF
            sProValida = sProValida & IIf(sProValida <> "", Chr(13), "") & RsTemp!tResumido & " (Precio Min.: " & Format(RsTemp!nPrecioMinimo, "##,##0.000") & " Precio Art: " & Format(RsTemp!nPrecio, "##,#0.000") & " Precio Max.: " & Format(RsTemp!nPrecioMaximo, "##,##0.000") & "), "
            RsTemp.MoveNext
        Loop
        Set RsTemp = Nothing
        nLen = Len(Trim(sProValida)) - 1
        sProValida = Left(Trim(sProValida), nLen)
        If nCount = 1 Then
            MsgBox "El articulo: " & Chr(13) & _
                   sProValida & Chr(13) & _
                   "no se encuentra dentro del Rango de su PRECIO DE COSTO minimo y/o maximo", vbInformation, sMensaje
        Else
            MsgBox "Los articulos: " & Chr(13) & _
                   sProValida & Chr(13) & _
                   "no se encuentran dentro del Rango de sus PRECIO DE COSTO minimos y/o maximos", vbInformation, sMensaje
        End If
        
        Screen.MousePointer = vbDefault
'        If Contraseña("07") = False Then
'            MsgBox "Clave no permitida", vbExclamation, sMensaje
'            Exit Sub
'
'        End If
        
'        lObserva = False
'        frmObservacion.Show vbModal
'        If lObserva = False Then
'            MsgBox "La contraseña no es válida...", vbCritical, sMensaje
'            Exit Sub
'        End If
'        Set mDocObservado = New clsDocumentoObservacion
'        mDocObservado.DocumentoObservacion frmDocumento.RsCabecera!tDocumento, frmDocumento.RsCabecera!tcodigoproveedor, Format(FechaServidor, "YYYY/mm/dd HH:mm:ss"), frmDocumento.RsCabecera!tCorrelativo, True, "I"
'
'        Set mDocObservado = Nothing
    End If
    

    
    Set RsTemp = Nothing
    If Calcular("Select lFecha as codigo from tParametro", CnAlmacen) = True Then
        fFechaProcesa = Format(Calcular("select getdate() as Codigo", CnAlmacen), "YYYY/mm/dd HH:mm:ss")
    Else
        fFechaProcesa = Format(Calcular("select getdate() as Codigo", CnAlmacen), "YYYY/mm/dd HH:mm:ss")
    End If
    
    
    FechaTransaccion = Format(DtpIngreso.value, "YYYY/mm/dd HH:mm") ' & " 00:00:00"
'    If lKardexFechaIngreso Then
'        'FechaTransaccion = Format(ComparaFecha(DtpIngreso), "YYYY/mm/dd HH:mm") '& " 00:00:00"
''        If lPermiteFecha = False Then Exit Sub
''        If (ValidaCierrePosterior(FechaTransaccion, CboArea.BoundText, "", CboTipoDocumento.BoundText)) > 0 Then
''            MsgBox "Existe un cierre de inventario posterior a la fecha de ingreso del documento. ", vbExclamation, sMensaje
''            Exit Sub
''        End If
'    End If
    
    
    Screen.MousePointer = vbHourglass
    Set mdDocumento = New ClsDocumento
    
    If lAuditoria Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "MDOCUMENTO", "INGRESO DE DOCUMENTOS", "02", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), LblCorrelativo, "", "tCorrelativo", "Correlativo", LblCorrelativo, "tEstadoDocumento", "Estado de Documento", "02", "tProcesa", "Usuario que Procesó", IIf(sUsuarioAutoriza <> "", sUsuarioAutoriza, sUsuario), "fProcesa", "Fecha de Proceso", FechaServidor())
    
        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    mdDocumento.CambiaEstadoDocumento LblCorrelativo, sUsuario, "02", Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss"), Format(DtpEmision.value, "YYYY/mm/dd"), 2, Format(FechaContable, "yyyy/mm/dd")
    Set mdDocumento = Nothing
    
    swEntrar = False
    'Asignar
    swEntrar = True
    
   '  Solamente se genera el movimiento de kardex, stocks y precios si el tipo de documento lo requiere...
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If CboArea.BoundText = "000" Then        'Almacen Central
        If Sw Then
            RsDtemporal.MoveFirst
            Do While Not RsDtemporal.EOF
                sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
                sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
                sValor = IIf(Not IsNull(RsDtemporal!nNeto), RsDtemporal!nNeto, "")
                xPUni = CDbl(IIf(Not IsNull(RsDtemporal!nPrecio), RsDtemporal!nPrecio, "0"))
                sAreaO = IIf(Not IsNull(RsDtemporal!tAreaOrigen), RsDtemporal!tAreaOrigen, "")
                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                    Consulta_Kardex "P"
                    CorrKD = CorrKD + 1
                    Isql = "Update tProducto Set " _
                           & "nPrecioCosto =" & xPUni & "," _
                           & "nPrecioPromedio =" & sPromedio & "," _
                           & "nStockActual=" & CDbl(sUltimo) + CDbl(sCant) & "," _
                           & "lreplica=1  " _
                           & "Where tCodigoProducto='" & sCodi & "'"
                    CnAlmacen.Execute Isql
                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
                    'Ingreso del Artículo...
                    Isql = "Insert Into mKardex (" _
                           & " tCodigoProducto,nCorrelativo,tTipoDocumento,tCodigoArea,tDocumento," _
                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
                           & " values ('" & sCodi & "', " _
                           & " '" & CorrKD & "', " _
                           & " '" & CboTipoDocumento.BoundText & "', " _
                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & xPUni * CDbl(sCant) & ", " _
                           & "" & 0 & ", " _
                           & "" & CDbl(sUltimo) + CDbl(sCant) & ", " _
                           & "" & CDbl(sUltimo) & ", " _
                           & "'01', " & IIf(sPromedio = 0, sCosto, sPromedio) & ", '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
                           & " '" & sUsuario & "', " _
                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
                    CnAlmacen.Execute Isql
                End If
                RsDtemporal.MoveNext
            Loop
        Else  'if sw
            RsDocumentoD.MoveFirst
            Do While Not RsDocumentoD.EOF
                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                sValor = IIf(Not IsNull(RsDocumentoD!nNeto), RsDocumentoD!nNeto, "")
                Select Case pais
                    Case "000", "002", "003", "004", "005"
                        xPUni = CDbl(IIf(Not IsNull(RsDocumentoD!nPrecio), RsDocumentoD!nPrecio, "0"))
                    Case "001"
                        xPUni = CDbl(IIf(Not IsNull(RsDocumentoD!nPrecio), RsDocumentoD!nPrecio, "0"))
                End Select
                sAreaO = IIf(Not IsNull(RsDocumentoD!tAreaOrigen), RsDocumentoD!tAreaOrigen, "")
              
                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                    Consulta_Kardex "P"
                    CorrKD = CorrKD + 1
                    
                    Isql = "Update tProducto Set " _
                           & "nPrecioCosto =" & xPUni & "," _
                           & "nPrecioPromedio =" & sPromedio & "," _
                           & "nStockActual=" & CDbl(sUltimo) + CDbl(sCant) & "," _
                           & "lreplica=1  " _
                           & "Where tCodigoProducto='" & sCodi & "'"
                    CnAlmacen.Execute Isql
                    
                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
                
                   'Ingreso del Artículo...
                    Isql = "Insert Into mKardex (" _
                           & " tCodigoProducto,nCorrelativo,tTipoDocumento,tCodigoArea,tDocumento," _
                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
                           & " values ('" & sCodi & "', " _
                           & " '" & CorrKD & "', " _
                           & " '" & CboTipoDocumento.BoundText & "', " _
                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
                           & "" & CDbl(sCant) & ", " _
                           & "" & xPUni * CDbl(sCant) & ", " _
                           & "" & 0 & ", " _
                           & "" & CDbl(sUltimo) + CDbl(sCant) & ", " _
                           & "" & CDbl(sUltimo) & ", " _
                           & "'01'," & " '" & IIf(sPromedio = 0, sCosto, sPromedio) & "', '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
                           & " '" & sUsuario & "', " _
                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
                    CnAlmacen.Execute Isql
                End If
                RsDocumentoD.MoveNext
            Loop
        End If
    Else 'Sub-Almacenes
'        sPrf = "RQ"
'        sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
'        If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
'            GeneraEncargado
'            lGeneraNS = True
'            Consulta_Kardex "P"
'            Isql = "Insert Into mRequerimiento (" _
'                   & "tRequerimiento,tCodigoArea,tResponsable," _
'                   & "tEstadoRequerimiento,tUsuario,fRegistro," _
'                   & "tProcesa,fProcesa,tDocumento,tCorrelativo,lReplica,fDiaContable,lPedido) " _
'                   & "Values ('" & Trim(sBRequerimiento) & "', " _
'                   & "'" & CboArea.BoundText & "', " & " '" & Trim(sEncargado) & "','02'," _
'                   & "'" & sUsuario & "', '" & Format(IIf(lKardexFechaIngreso, FechaTransaccion, fFechaProcesa), "YYYY/mm/dd HH:mm:ss") & "', " _
'                   & "'" & sUsuario & "', '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
'                   & "'" & Trim(LblPDoc & TxtSerie & TxtCorrelativo) & "'," _
'                   & "'" & LblCorrelativo & "' , 1,convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),0)"
'            CnAlmacen.Execute Isql
'
'            '~~~ Graba el Requerimiento automático en la cabecera de Documentos ~~~
'            Isql = "Update mDocumento SET " _
'                   & " tRequerimiento ='" & Trim(sBRequerimiento) & "'" _
'                   & " ,lreplica=1 WHERE tCorrelativo='" & LblCorrelativo & "'"
'            CnAlmacen.Execute Isql
'        End If
'
'        If Sw Then
'            RsDtemporal.MoveFirst
'            Do While Not RsDtemporal.EOF
'                sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
'                sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
'                sValor = IIf(Not IsNull(RsDtemporal!nNeto), RsDtemporal!nNeto, "")
'                nDItem = RsDtemporal!nItem
'                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
'
'                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
'                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
'
'                    Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, (((b.nstockactual * b.nPrecioPromedio)+ " _
'                           & " (a.ncantidad * a.nprecio))/CASE ((a.ncantidad + b.nstockactual) WHEN 0 THEN 1 ELSE (a.ncantidad + b.nstockactual) END) as PrecioPromedio, (a.ncantidad + b.nstockactual) as stockact,b.nPrecioCosto from ddocumento a," _
'                           & " tproducto b where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                    Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
'                    If Not RsPromedio.EOF Then
'                        RsPromedio.MoveLast
'                        sPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
'                    End If
'
'                    If sUltRequerimiento = "" Then sUltRequerimiento = 0
'                    Isql = "Insert Into dRequerimiento (" _
'                           & "tRequerimiento,tCodigoProducto," _
'                           & " nCantidadPedida,nAtendido,nPendiente,nStockActual,nPrecioPromedio,nItem)" _
'                           & " Values ('" & Trim(sBRequerimiento) & "', " _
'                           & " '" & sCodi & "', " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & 0 & ", " _
'                           & "" & CDbl(sUltRequerimiento) + CDbl(sCant) & ", " _
'                           & "" & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & "," & nDItem & ")"
'                    CnAlmacen.Execute Isql
'
'                    If Calcular("select lPorcionablePL as codigo from tParametro", CnAlmacen) = True Then
'                        Set clsPorciona = New clsPorcionamiento
'                        Isql = "SELECT tArticuloPorcionado.tcodigoarticuloRes,dbo.tArticuloPorcionado.tcodigoPorcionable,dbo.tArticuloPorcionado.tcodigoPorcionado, dbo.tArticuloPorcionado.nPorcentajeMerma, dbo.tArticuloPorcionado.nEquivalencia, dbo.TPRODUCTO.tDetallado FROM dbo.tArticuloPorcionado INNER JOIN dbo.TPRODUCTO ON dbo.tArticuloPorcionado.tCodigoPorcionado = dbo.TPRODUCTO.tCodigoProducto WHERE  (dbo.tArticuloPorcionado.lPorcAutomatico = 1) and tcodigoporcionable='" & sCodi & "'"
'                        Set rsPlantilla = Lib.OpenRecordset(Isql, CnAlmacen)
'                        If rsPlantilla.RecordCount > 0 Then
'                            rsPlantilla.MoveFirst
'                            sCodPorciona = clsPorciona.InsUpdPorciona("", sCodi, CboArea.BoundText, sUsuario, CDbl(sCant), CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), "01", "NULL", "NULL", Trim(sBRequerimiento), "I", True, Format(IIf(lKardexFechaIngreso, FechaTransaccion, fFechaProcesa), "YYYY/mm/dd HH:mm:ss"))
'                            Isql = "Insert into dPorcionamiento (tCodigoPorcionamiento, tCodigoProducto, tDetallado, tCantidadProducto, tEquivalencia, nPProm)  " & _
'                                   " values ('" & sCodPorciona & "','" & rsPlantilla!tCodigoPorcionado & "','" & rsPlantilla!tDetallado & "'," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & "," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & ", 0)"
'                            CnAlmacen.Execute Isql
'                        End If
'                        Set clsPorciona = Nothing
'                    End If
'                End If
'                RsDtemporal.MoveNext
'            Loop
'        Else
'            RsDocumentoD.MoveFirst
'            Do While Not RsDocumentoD.EOF
'                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
'                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
'                sValor = IIf(Not IsNull(RsDocumentoD!nNeto), RsDocumentoD!nNeto, "")
'                nDItem = RsDocumentoD!nItem
'                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
'                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
'                    sCosto = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
'
'                    Isql = "Select a.tcodigoproducto, a.nprecio, a.ncantidad, b.nstockactual, b.nPrecioPromedio, (((b.nstockactual * b.nPrecioPromedio)+ " _
'                           & " (a.ncantidad * a.nprecio))/ (CASE a.ncantidad + b.nstockactual WHEN 0 THEN 1 ELSE a.ncantidad + b.nstockactual END)) as PrecioPromedio, (a.ncantidad + b.nstockactual) as stockact,b.nPrecioCosto from ddocumento a," _
'                           & " tproducto b where a.tcodigoproducto = b.tcodigoproducto And b.tCodigoProducto='" & sCodi & "' And a.tCorrelativo='" & LblCorrelativo & "'"
'                    Set RsPromedio = Lib.OpenRecordset(Isql, CnAlmacen)
'                    If Not RsPromedio.EOF Then
'                        RsPromedio.MoveLast
'                        sPromedio = IIf(IsNull(RsPromedio!PrecioPromedio), 0, RsPromedio!PrecioPromedio)
'                    End If
'
'                    If sUltRequerimiento = "" Then sUltRequerimiento = 0
'                    Isql = "Insert Into dRequerimiento (" _
'                           & "tRequerimiento,tCodigoProducto," _
'                           & " nCantidadPedida,nAtendido,nPendiente,nStockActual,nPrecioPromedio,nItem)" _
'                           & " Values ('" & Trim(sBRequerimiento) & "', " _
'                           & " '" & sCodi & "', " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & 0 & ", " _
'                           & "" & CDbl(sUltRequerimiento) + CDbl(sCant) & ", " _
'                           & "" & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & "," & nDItem & ")"
'                    CnAlmacen.Execute Isql
'
'                    If Calcular("select lPorcionablePL as codigo from tParametro", CnAlmacen) = True Then
'                        Set clsPorciona = New clsPorcionamiento
'                        Isql = "SELECT tArticuloPorcionado.tcodigoarticuloRes,dbo.tArticuloPorcionado.tcodigoPorcionable,dbo.tArticuloPorcionado.tcodigoPorcionado, dbo.tArticuloPorcionado.nPorcentajeMerma, dbo.tArticuloPorcionado.nEquivalencia, dbo.TPRODUCTO.tDetallado FROM dbo.tArticuloPorcionado INNER JOIN dbo.TPRODUCTO ON dbo.tArticuloPorcionado.tCodigoPorcionado = dbo.TPRODUCTO.tCodigoProducto WHERE  (dbo.tArticuloPorcionado.lPorcAutomatico = 1) and tcodigoporcionable='" & sCodi & "'"
'                        Set rsPlantilla = Lib.OpenRecordset(Isql, CnAlmacen)
'                        If rsPlantilla.RecordCount > 0 Then
'                            rsPlantilla.MoveFirst
'                            sCodPorciona = clsPorciona.InsUpdPorciona("", sCodi, CboArea.BoundText, sUsuario, CDbl(sCant), CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100), "01", "NULL", "NULL", Trim(sBRequerimiento), "I", True, Format(IIf(lKardexFechaIngreso, FechaTransaccion, fFechaProcesa), "YYYY/mm/dd HH:mm:ss"))
'                            Isql = "Insert into dPorcionamiento (tCodigoPorcionamiento, tCodigoProducto, tDetallado, tCantidadProducto, tEquivalencia, nPProm)  " & _
'                                   " values ('" & sCodPorciona & "','" & rsPlantilla!tCodigoPorcionado & "','" & rsPlantilla!tDetallado & "'," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & "," & CDbl(sCant) - (CDbl(sCant) * rsPlantilla!nPorcentajeMerma / 100) & ", 0)"
'                            CnAlmacen.Execute Isql
'                        End If
'                        Set clsPorciona = Nothing
'                    End If
'                End If
'                RsDocumentoD.MoveNext
'            Loop
'        End If
'
'        If Sw Then
'            RsDtemporal.MoveFirst
'            Do While Not RsDtemporal.EOF
'                sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
'                sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
'                xPUni = IIf(Not IsNull(RsDtemporal!nPrecio), RsDtemporal!nPrecio, "")
'                sAreaO = IIf(Not IsNull(RsDocumentoD!tAreaOrigen), RsDocumentoD!tAreaOrigen, "")
'                sValor = IIf(Not IsNull(RsDtemporal!nNeto), RsDtemporal!nNeto, "")
'                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
'                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
'                    lGeneraNS = False
'                    Consulta_Kardex "P"
'                    CorrKD = CorrKD + 1
'
'                    'Ingreso
'                    Isql = "Insert Into mKardex (" _
'                           & "tCodigoProducto,nCorrelativo,tTipoDocumento,linterno,tCodigoArea,tDocumento," _
'                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
'                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
'                           & " Values ('" & sCodi & "', " _
'                           & " '" & CorrKD & "','" & CboTipoDocumento.BoundText & "','1'," _
'                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
'                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & CDbl(xPUni) * CDbl(sCant) & ", " _
'                           & "" & 0 & ", " & " " & CDbl(sAnteriorKD) + CDbl(sCant) & ", " _
'                           & "" & CDbl(sAnteriorKD) & ", " _
'                           & "'01'," & " " & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & ", '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
'                           & " '" & sUsuario & "', " _
'                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
'                    CnAlmacen.Execute Isql
'                    Isql = "update tProducto set nStockActual=isnull(nStockActual,0)+" & CDbl(sCant) & " ,lreplica=1  where tCodigoProducto='" & sCodi & "'"
'                    CnAlmacen.Execute Isql
'                    GeneraSalida     'Genera la salida del item antes de pasar al otro...
'                End If
'
'                RsDtemporal.MoveNext
'            Loop
'        Else                      'Se está modificando...
'            RsDocumentoD.MoveFirst
'            Do While Not RsDocumentoD.EOF
'                sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
'                sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
'                xPUni = IIf(Not IsNull(RsDocumentoD!nPrecio), RsDocumentoD!nPrecio, "")
'                sAreaO = IIf(Not IsNull(RsDocumentoD!tAreaOrigen), RsDocumentoD!tAreaOrigen, "")
'                sValor = IIf(Not IsNull(RsDocumentoD!nNeto), RsDocumentoD!nNeto, "")
'                sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
'
'                If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
'                    lGeneraNS = False
'                    Consulta_Kardex "P"
'                    CorrKD = CorrKD + 1
'
'                    'Ingreso
'                    Isql = "Insert Into mKardex (" _
'                           & "tCodigoProducto,nCorrelativo,tTipoDocumento,linterno,tCodigoArea,tDocumento," _
'                           & " nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
'                           & " nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable,ftransaccion)" _
'                           & " Values ('" & sCodi & "', " _
'                           & " '" & CorrKD & "','" & CboTipoDocumento.BoundText & "','1'," _
'                           & " '" & IIf(sAreaO = "", CboArea.BoundText, sAreaO) & "', " _
'                           & " '" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "', " _
'                           & "" & CDbl(sCant) & ", " _
'                           & "" & CDbl(xPUni) * CDbl(sCant) & ", " _
'                           & "" & 0 & ", " & " " & CDbl(sAnteriorKD) + CDbl(sCant) & ", " _
'                           & "" & CDbl(sAnteriorKD) & ", " _
'                           & "'01'," & " " & IIf(CDbl(sPromedio) = 0, CDbl(sCosto), CDbl(sPromedio)) & ", '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
'                           & " '" & sUsuario & "', " _
'                           & " " & xPUni & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
'                    CnAlmacen.Execute Isql
'                    'DETP se agrego el precio de promedio 10/08/2007
'                    Isql = "update tProducto set nStockActual=isnull(nStockActual,0)+" & CDbl(sCant) & "" & IIf(CDbl(sPromedio) = 0, "", " , nPrecioPromedio =" & sPromedio & "") & ",lreplica=1  where tCodigoProducto='" & sCodi & "'"
'                    CnAlmacen.Execute Isql
'                    GeneraSalida     'Genera la salida del item antes de pasar al otro...
'                End If
'
'                RsDocumentoD.MoveNext
'            Loop
'        End If
    End If
      
   
   

   
'    cHistorial.FunInsertatHistorial "001", "007", "", LblCorrelativo, "", "", "", "", sUsuario, sUsuarioAutoriza, "Procesó el documento " & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & ""
   
    CmdProcesar.Caption = "Desprocesar"

    Set mdDocumento = New ClsDocumento
    mdDocumento.HistorialCompra LblCorrelativo, "I"
    Set mdDocumento = Nothing

    Screen.MousePointer = vbDefault
    MsgBox "Información procesada...", vbInformation, sMensaje
    
'    Dim iRow As Integer
'    frmDocumento.RsCabecera.Requery
'    If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.MoveFirst
'    frmDocumento.RsCabecera.Find "tCorrelativo='" & LblCorrelativo & "'"
    
    swEntrar = False
    'Asignar
    swEntrar = True
    
    ActivaAdicion False
    ControlEstado
    grdGrilla.MoveFirst
    sUsuarioAutoriza = ""
End Sub
Sub GeneraSalida()
    Dim FechaContable As Date
    Dim oDiaContable As New clsDiaContable
    Dim dStock As Double
    
    FechaContable = oDiaContable.obtieneDiaContable
    
    dStock = sAnteriorKD

    'Genero Registro Nota Salida en Almacen Central
    sPrf = "RQ"
        
    lGeneraNS = False
    CorrKD = CorrKD + 1
                           
    '~~~~~~ REGISTRA LA SALIDA ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Modificación CLR 08/09/2011:Se quitaron todas las llamadas al procedimeitno Consulta_Kardex ya que ocacionaba el recalculo del equivalente en moneda nacional para el precio de costo.
    CorrKD = CorrKD + 1
         
    sPrf = "RQ"
    sAnteriorKD = Calcular("Select nStockActual As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen)
    
    Isql = "Insert Into mKardex (" _
         & "tCodigoProducto,nCorrelativo,tTipoDocumento,linterno,tCodigoArea,tDocumento," _
         & "nIngreso,nValor,nSalida,nStockActual,nStockUltimo,tEstadoKardex," _
         & "nPrecioPromedio,fRegistro,tUsuario,nPrecioCosto,tCorrelativo,fDiaContable, ftransaccion) " _
         & "Values ('" & sCodi & "'," & "'" & CorrKD & "','99','1'," _
         & "'" & CboArea.BoundText & "'," _
         & "'" & Trim(sPrf) & "-" & Trim(sBRequerimiento) & "'," _
         & "" & 0 & "," & "" & CDbl(sPromedio) * CDbl(sCant) & "," _
         & "" & CDbl(sCant) & "," _
         & "" & CDbl(sAnteriorKD) - CDbl(sCant) & "," _
         & "" & CDbl(sAnteriorKD) & "," & "'01'," & "" & IIf(sPromedio = 0, CDbl(sCosto), CDbl(sPromedio)) & ", '" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', " _
         & "'" & (sUsuario) & "'," & "" & CDbl(sPromedio) & ", '" & LblCorrelativo & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')"
    CnAlmacen.Execute Isql
         
    Isql = "Update tProducto Set " _
         & "nPrecioCosto=" & xPUni & "," _
         & "nPrecioPromedio=" & sPromedio & "," _
         & "nStockActual =" & CDbl(sAnteriorKD) - CDbl(sCant) & ", " _
         & "lreplica=1  " _
         & "Where tCodigoProducto='" & sCodi & "'"
    CnAlmacen.Execute Isql
         
    Dim xPrecioPromedio As Double
    xPrecioPromedio = sPromedio

    'Genero el Proceso en tSubStock (Modificación o Ingreso)
'    Consulta_Kardex
    
    CorrSTK = CorrSTK + 1
    If sCodi = xCodigo And xArea = CboArea.BoundText Then
       Isql = "Update tSubStock Set " _
            & "nCorrelativo ='" & CorrSTK & "', " _
            & "nStockActual =" & CDbl(sCant) + CDbl(sSub_Stock) & ",lActivo=1 " _
            & "Where tCodigoProducto='" & sCodi & "' And " _
            & "tCodigoSubArea ='" & CboArea.BoundText & "'"
       CnAlmacen.Execute Isql
    Else
      'Preguntar si esta sino Adicionar
       Isql = "Insert Into tSubStock (" _
            & "tCodigoProducto,tCodigoSubArea,nCorrelativo,nStockActual,lActivo) " _
            & "Values ('" & sCodi & "', " _
            & "'" & CboArea.BoundText & "', " _
            & "'" & CorrSTK & "', " _
            & "" & CDbl(sCant) & ",1)"
       CnAlmacen.Execute Isql
    End If
    
    'Genero la Nota Salida al Sub-Almacen (Ingresa al área destino)
    lGeneraNS = False    'NUEVO
    CorrSKD = CorrSKD + 1
    sPrf = "RQ"
    If sAnteriorSKD = "" Then sAnteriorSKD = 0
    
    Isql = "Insert Into mSubKardex (tCodigoSubArea,tCodigoProducto,     nCorrelativo,tTipoDocumento,tCodigoArea,                                        tDocumento,           nIngreso,                                     nValor,  nSalida,                            nStockActual,              nStockUltimo,tEstadoKardex,              nPrecioPromedio,                  nPrecioCosto,                                            fRegistro,           tUsuario,                                                        fDiaContable,ftransaccion)" & _
              "Values ('" & CboArea.BoundText & "','" & sCodi & "','" & CorrSKD & "',          '99',      '000','" & Trim(sPrf) & "-" & Trim(sBRequerimiento) & "'," & CDbl(sCant) & "," & CDbl(sCant) * CDbl(xPrecioPromedio) & "," & 0 & "," & CDbl(sCant) + CDbl(sAnteriorSKD) & "," & CDbl(sAnteriorSKD) & ",         '01'," & CDbl(xPrecioPromedio) & "," & CDbl(xPrecioPromedio) & ",'" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "', '" & sUsuario & "',convert(smalldatetime,'" & Format(FechaContable, "yyyy/mm/dd") & "'),'" & FechaTransaccion & "')   "
     CnAlmacen.Execute Isql
    'Salida Mermas Fijas
    CorrSKD = CorrSKD + 1

End Sub
Public Sub GeneraEncargado()
'   Isql = "Select Min(vGrupoResponsable.tCodigo) AS Codigo " _
'        & " From vGrupoArea Left Outer Join vGrupoResponsable ON vGrupoArea.Codigo = vGrupoResponsable.tValor" _
'        & " Where vGrupoResponsable.tValor ='" & cboArea.BoundText & "'"
    Isql = "sp_TraeGeneraEncargado '" & CboArea.BoundText & "'"
    Set RsGEncargado = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsGEncargado.EOF Then
       RsGEncargado.MoveFirst
       sEncargado = IIf(IsNull(RsGEncargado!codigo), "", RsGEncargado!codigo)
    End If
End Sub

Sub FechaEmision()
    If DtpEmision > CDate(Format(FechaServidor, "dd/mm/yyyy")) Then MsgBox "Error en la Fecha de Emisión...", vbCritical, sMensaje: Exit Sub
    If DtpEmision > DtpIngreso Then MsgBox "Error en las Fechas de Emisión o Ingreso...", vbCritical, sMensaje: Exit Sub
    If DtpIngreso < DtpEmision Then MsgBox "Error en las Fechas de Emisión o Ingreso...", vbCritical, sMensaje: Exit Sub
    Dim RsFechas As ADODB.Recordset
    Isql = "SELECT [fFecha],[Fregistro],[nVenta] From tTipoCambio Where fFecha='" & Format(DtpEmision.value, "yyyyMMdd") & "'"
    Set RsFechas = Lib.OpenRecordset(Isql, CnAlmacen)
    
    If Not RsFechas.EOF Then
        LblTipoCambio = Format(IIf(IsNull(RsFechas!nVenta), "", RsFechas!nVenta), "##0.000")
    Else
        MsgBox "Es necesario que ingrese ahora" & vbCrLf & "el Tipo de Cambio en almacen para continuar...", vbInformation, "Almacen"
        DtpEmision.Enabled = True
        'sFormulario = "frmDocumentoCompra"
'        With FrmCreaTipoCambio
'            .DtpCambio = DtpEmision
'            .Show vbModal
'        End With
        'sFormulario = ""
        
        swEntrar = False
        
        Isql = "SELECT [fFecha],[Fregistro],[nVenta] From tTipoCambio Where fFecha='" & Format(DtpEmision.value, "MM/dd/yyyy") & "'"
        Set RsFechas = Lib.OpenRecordset(Isql, CnAlmacen)
        If RsFechas.RecordCount > 0 Then
            LblTipoCambio = Format(IIf(IsNull(RsFechas!nVenta), "", RsFechas!nVenta), "##0.000")
        Else
            LblTipoCambio = 0
        End If
    End If
    Set RsFechas = Nothing
End Sub

Private Sub DtpEmision_Change()
    If cmdOpcion(3).Visible Then
        cmdOpcion(4).Enabled = True
        xDtpEmision = True
    End If
End Sub

Private Sub DtpEmision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

'Private Sub DtpEmision_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        'txtRuc.SetFocus
'        DtpIngreso.SetFocus
''    Else
''        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
''            MsgBox "Ingrese únicamente números...", vbInformation, sMensaje
''            KeyAscii = 0
''            TxtCorrelativo.SetFocus
''        End If
'    End If
'End Sub

Private Sub DtpEmision_LostFocus()
    If DtpEmision.value > CDate(Format(FechaServidor, "dd/mm/yyyy")) Then
        MsgBox "Error en la Fecha de Emisión...", vbCritical, sMensaje
        If Sw Then DtpEmision.value = CDate(Format(FechaServidor, "dd/mm/yyyy"))
        DtpEmision.SetFocus
    End If
    If DtpIngreso < DtpEmision Then
        MsgBox "Error en la Fecha de Emisión o Ingreso ...", vbCritical, sMensaje
        If Sw Then
            DtpEmision.value = CDate(Format(FechaServidor, "dd/mm/yyyy"))
        End If
    End If
    If DtpEmision.value < CDate(Format(FechaServidor, "dd/mm/yyyy")) Then
        FechaEmision
    Else
        Exit Sub
    End If
End Sub



Private Sub DtpIngreso_Change()
    
    If cmdOpcion(3).Visible Then
        cmdOpcion(4).Enabled = True
        xDtpIngreso = True
    End If
End Sub

Private Sub DtpIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpIngreso_LostFocus()
    If DtpIngreso.value > CDate(Format(FechaServidor, "dd/mm/yyyy")) Then
        MsgBox "Error en la Fecha de Ingreso...", vbCritical, sMensaje
        DtpIngreso.value = CDate(Format(FechaServidor, "dd/mm/yyyy"))
        DtpIngreso.SetFocus
    End If
    If DtpIngreso < DtpEmision Then
        MsgBox "Error en la Fecha de Emisión o Ingreso ...", vbCritical, sMensaje
'        DtpEmision.value = Format(FechaServidor, "dd/mm/yyyy")
        DtpIngreso.value = CDate(Format(FechaServidor, "dd/mm/yyyy"))
    End If
    'TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
    'If pais <> "002" Then
        dtpAnio.value = Format(DtpIngreso.value, "dd/mm/yyyy")
        LlenaMesContable
    'End If
End Sub

Private Sub Form_Load()
    Centrar Me
    Sw = True
    lModifica = False
    
    lLoad = True
    lValidaProducto = True
    'lGeneraNS = False
    lActualiza = False

    'sForm = "frmDocumentoCompra"
    tPagoRapidoTipoPago = Calcular("select isnull(tPagoRapidoTipoPago,'') as codigo from tParametro", CnAlmacen)
    '|||Jesus|||
    lMotivoConsumo = Calcular("select isnull(lMotivoConsumo,0) as codigo from tParametro", CnAlmacen)
    '|||||||||||
    
    If (Trim(tPagoRapidoTipoPago) = "") Then
        tPagoRapidoTipoPago = Calcular("select top 1 Codigo from vFormaPago where tResumido like '%EFECTIVO%' ", CnAlmacen)
    End If
    
    
    grdGrilla.Height = 4065     '<----- OJO ----- No variar
    
    DtpEmision = Format(FechaServidor, "dd/mm/yyyy")
    'DtpFPrograma = Format(FechaServidor, "dd/mm/yyyy")
    DtpIngreso = Format(FechaServidor, "dd/mm/yyyy")
    'dtpFechCaducidad = Format(FechaServidor, "dd/mm/yyyy")
    dtpAnio = Format(FechaServidor, "dd/mm/YYYY")
    
    'dtpContable = Format(FechaServidor, "dd/mm/yyyy")
    
    'TxtNombre.Enabled = True
    txtTransporte.Enabled = True
    'TxtProveedor.Enabled = True
    'TxtRUC.Enabled = True
    
    'cmdDescuento.Enabled = True     'False
    CmdObservacion.Enabled = False
    
    LblImpt1 = sImpuesto1
    LblImpt2 = sImpuesto2
    LblImpt3 = sImpuesto3
    
    'chkDetraccion.Caption = LlenaLey("C") & " :"
'    lblLeyD.Caption = LlenaLey("D")
'    If chkDetraccion.Caption = " :" Then
'        chkDetraccion.Visible = False
'    End If
    swEntrar = False
    sFlag = False
    iTBox = 0
    
    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "setiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    
    LlenaCombos
    LlenaMesContable
    ActivaAdicion True
    
    'solo para el tema de bolivia
     Select Case pais
        Case "000", "004", "005"
            lblCodigoControl.Visible = False
            'txtCodigoControl.Visible = False
            lblCodigoAutori.Visible = False
            'txtAutorizacion.Visible = False
            lblFechCaducidad.Visible = False
            'dtpFechCaducidad.Visible = False
            TxtICE.Visible = False
            LblICE.Visible = False
            Shape11.Visible = False
            Label15(1).Visible = False
        Case "001"
            lblCodigoControl.Visible = True
            'txtCodigoControl.Visible = True
            lblCodigoAutori.Visible = True
            'txtAutorizacion.Visible = True
            lblFechCaducidad.Visible = False
            'dtpFechCaducidad.Visible = False
            TxtICE.Visible = False
            LblICE.Visible = True
            Shape11.Visible = True
            Label15(1).Visible = True
            LblImpt3.Visible = False
            LblTotImp3.Visible = False
            Shape3.Visible = False
        Case "002"
            lblCodigoControl.Visible = False
            'txtCodigoControl.Visible = False
            lblCodigoAutori.Visible = True
            'txtAutorizacion.Visible = True
            lblFechCaducidad.Visible = True
            'dtpFechCaducidad.Visible = True
            TxtICE.Visible = False
            LblICE.Visible = False
            Shape11.Visible = False
            Label15(1).Visible = False
            
    End Select
    
'    If frmDocumento.RsCabecera.RecordCount > 0 Then
'        If Not frmDocumento.RsCabecera.EOF Then
'            If UCase(frmDocumento.RsCabecera!CabEstado) = "PROCESADO" Then
'                CmdDescuento.Enabled = False
'                cmdOpcion(10).Enabled = False
'                CboReferencia.Enabled = False
'            Else
'                CmdDescuento.Enabled = True
'                'Bolivia
'                Select Case pais
'                    Case "000", "005": cmdOpcion(10).Enabled = True
'                    Case "001": cmdOpcion(10).Enabled = False
'                    Case "002": cmdOpcion(10).Enabled = True
'                End Select
'            End If
'        Else
'            CmdDescuento.Enabled = False
'            cmdOpcion(10).Enabled = False
'            CboReferencia.Enabled = False
'        End If
'    End If
    Select Case pais
    Case "000", "002", "004", "003", "005"
        ConfGrilla 10, grdGrilla, "Código", 2, "tCodigoProducto", 890, 2, 0, "", _
                             "Artículo", 2, "Producto", 3880, 0, 0, "", _
                             sImpuesto1, 1, "nImpuesto1", 855, 1, 0, "###,###,##0.00", _
                             sImpuesto2, 1, "nImpuesto2", 855, 1, 0, "###,###,##0.00", _
                             sImpuesto3, 1, "nImpuesto3", 855, 1, 0, "###,###,##0.00", _
                             "Unidad Kardex", 2, "UnidadMedida", 1420, 2, 0, "", _
                             "Cantidad", 1, "nCantidad", 850, 1, 0, "0.00", _
                             "Precio", 1, "nPrecio", 750, 1, 0, "###,###,##0.000", _
                             "Precio Total", 1, "nNeto", 1100, 1, 0, "###,###,##0.00", _
                             "Origen", 3, "tAreaOrigen", 50, 1, 0, ""
        grdGrilla.Columns(9).Visible = False
     Case "001"
         ConfGrilla 10, grdGrilla, "Código", 2, "tCodigoProducto", 890, 2, 0, "", _
                             "Artículo", 2, "Producto", 3880, 0, 0, "", _
                             sImpuesto1, 1, "nImpuesto1", 855, 1, 0, "###,###,##0.00", _
                             sImpuesto2, 1, "nImpuesto2", 855, 1, 0, "###,###,##0.00", _
                             sImpuesto3, 1, "nImpuesto3", 855, 1, 0, "###,###,##0.00", _
                             "Unidad Kardex", 2, "UnidadMedida", 1420, 2, 0, "", _
                             "Cantidad", 1, "nCantidad", 850, 1, 0, "0.00", _
                             "Precio", 1, "nPrecio", 750, 1, 0, "###,###,##0.000", _
                             "Costo Total", 1, "nNeto", 1100, 1, 0, "###,###,##0.00", _
                             "Origen", 3, "tAreaOrigen", 50, 1, 0, ""
        grdGrilla.Columns(9).Visible = False
    End Select
    'Se le Agrego a nPrecio un decimal mas
    If Sw Then
        ActivarBotones False
        Blanquear Me
        'CboReferencia.Enabled = True
'        TxtSerieRef.Enabled = False
'        TxtCorrelativoRef.Enabled = False
        txtFechaProceso.Enabled = False
        
        'SwTmp = True
        swEntrar = True
        
        CargaTemporal
        CboArea.BoundText = "000"        'Almacen Central
        txtTransporte.Text = "0.00"
        CboMoneda.BoundText = "01"
        CboTipoIngreso.BoundText = "M"
        cboOperacion.BoundText = Calcular("select tCodigo as Codigo from tTabla where tValor='G' and tTabla='TOPERACION' and lActivo=1", CnAlmacen)
        cboOperacion.Enabled = False
        'CboReferencia.Enabled = False
        
        'ChkPrograma.Enabled = True
        'TxtDiasCredito.backColor = &HFFFFFF          'Habilita
        'TxtDiasCredito.Enabled = True
        'DtpFPrograma.Enabled = True
        'dtcProgramaPago.Enabled = True
        'TxtDiasCredito = Calcular("Select nCredito As Codigo From tProveedor Where tCodigoProveedor='" & TxtProveedor & "'", CnAlmacen)
       
       
     
        lDesc = ""
        
        LblTotImp1.ToolTipText = ""
        LblTotImp2.ToolTipText = ""
        LblTotImp3.ToolTipText = ""
        lblRedondeo.ToolTipText = ""
        LblTotal.ToolTipText = ""

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If CDbl(LblTotImp1) = 0 And LblImpt1 = "" Then
            Shape1.Visible = False
            LblTotImp1.Visible = False
        Else
            Shape1.Visible = True
            LblTotImp1.Visible = True
        End If
        If CDbl(LblTotImp2) = 0 And LblImpt2 = "" Then
            Shape2.Visible = False
            LblTotImp2.Visible = False
        Else
            Shape2.Visible = True
            LblTotImp2.Visible = True
        End If
        If CDbl(LblTotImp3) = 0 And LblImpt3 = "" Then
            Shape3.Visible = False
            LblTotImp3.Visible = False
        Else
            Shape3.Visible = True
            LblTotImp3.Visible = True
        End If
        
    Else

        swEntrar = False
        'Asignar
        swEntrar = True
        
        LlenaGrilla
        ActivarBotones True
        EstadoCabecera False
        ControlEstado
        
        'SwTmp = False
               
        TxtGlosaDocumento.backColor = &HF0EADB           '&H80000005
        'TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
            LblTotImp1.ToolTipText = ""
            LblTotImp2.ToolTipText = ""
            LblTotImp3.ToolTipText = ""
            lblRedondeo.ToolTipText = ""
            LblTotal.ToolTipText = ""
        '--Se cambio para la modificacion
        CboTipoIngreso.Enabled = False
            If CboTipoDocumento.BoundText = "01" Or CboTipoDocumento.BoundText = "12" Then
                cboOperacion.Enabled = True
            Else
                cboOperacion.Enabled = False
            End If
    End If
    
    '------------------ Se actualizo 16/01/06
    sPromedio = 0
    sUltPromedio = 0
    sUltimo = 0
    sAnteriorKD = 0
    sAnteriorSKD = 0
    sPCosto = 0
    sSub_Stock = 0
    sUltRequerimiento = ""
    lsAnula = ""
    CorrKD = 0
    CorrSKD = 0
    CorrSTK = 0
    sBRequerimiento = ""
    
    'PreparaTipoCambio
    'cmdTexto.Caption = "Registro " & IIf(frmDocumento.RsCabecera.RecordCount = 0, 0, frmDocumento.RsCabecera.AbsolutePosition) & " de " & frmDocumento.RsCabecera.RecordCount
    
'    If Not Sw Then
'        If frmDocumento.RsCabecera.RecordCount > 0 Then
'            If Not frmDocumento.RsCabecera.EOF Then
'                If IsNull(frmDocumento.RsCabecera!lDescuento) Then
'                    lDesc = ""
'                Else
'                    If frmDocumento.RsCabecera!lDescuento Then
'                        lDesc = "G"
'                    Else
'                        lDesc = "I"
'                    End If
'                End If
'            Else
'                lDesc = ""
'            End If
'        End If
'    End If
    
    CboTipoDocumento_LostFocus
        
   
    
    'FinModifica
    Tiempo = 0          '30000      'Inactividad por 30 segundos...
    
    Screen.MousePointer = vbDefault

    'Para ver se prenda el boton de Ordenes de Compra
    'cmdOpcion(6).Visible = False
    
    cmdPago.Visible = Calcular("select top 1 lPagoCajaChica as Codigo from tParametro", CnAlmacen)
    lExporta = Calcular("select top 1 lExportar as Codigo from tParametro", CnAlmacen)

    lblMontoControl.Visible = False
    txtMontoControl.Visible = False
        
    Select Case pais
        Case "000", "002", "004", "003", "005"
        Case "001"
            lblCodigoControl.Enabled = False
    End Select
    
        'dtcProgramaPago.Visible = True
        'Label17.Visible = True
        'TxtDiasCredito.Visible = False
        'DtpFPrograma.Visible = False
        'ChkPrograma.Visible = False
    
         lblLibroCompra.Visible = False
        dtcLibroCompra.Visible = False
    
    If pais = "002" Then
        label(6).Visible = True
        txtTransporte.Visible = True
        LabelMesContable.Visible = False
        dtpAnio.Visible = False
        CmbMes.Visible = False
        Shape7.Visible = False
    ElseIf pais = "000" Or pais = "004" Or pais = "005" Then
            LabelMesContable.Visible = False
            dtpAnio.Visible = False
            CmbMes.Visible = False
            Shape7.Visible = False
            
        label(6).Visible = False
        txtTransporte.Visible = False
        txtTransporte.Text = "0.00"
    Else
        LabelMesContable.Visible = False
        dtpAnio.Visible = False
        CmbMes.Visible = False
        Shape7.Visible = False
        label(6).Visible = False
        txtTransporte.Visible = False
        txtTransporte.Text = "0.00"
    End If

    lActualizaImpuesto = False
    
    Isql = "Select [fFecha], [nCompra], [nVenta] From tTipoCambio " _
           & " Where fFecha='" & Format(FechaServidor, "yyyyMMdd") & "'"
    Set RsCambio = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsCambio.EOF Then
        RsCambio.MoveLast
        LblTipoCambio.Caption = Format(RsCambio!nVenta, "##0.000")
    Else
        LblTipoCambio.Caption = "0.000"
    End If
    
    
End Sub

Sub CargaTemporal()
    sDtemporal = dbTemporalAlmacen(sUsuario, 40, "tDocumento", "nvarchar(15)", _
                                          "tCodigoProveedor", "nVarChar(80)", "nItem", "nVarChar(3)  ", _
                                          "tCodigoProducto", "nVarChar(7)  ", "tCodigoBarra", "nVarChar(20)  ", _
                                          "tUnidadEntrada", "nVarChar(3)  ", "nCantidad", "Float", _
                                          "nPrecio", "Float", "nPVenta", "Float", "nNeto", "Float", _
                                          "nDescuento", "Float", "nRecargo", "Float", "nImpuesto1", "Float", _
                                          "nImpuesto2", "Float", "nImpuesto3", "Float", "nPorcentaje1", "Float", _
                                          "nPorcentaje2", "Float", "nPorcentaje3", "Float", "nTotal", "Float", _
                                          "tCorrelativo", "nVarChar(11)  ", "tEstadoDocumento", "nVarChar(2)  ", _
                                          "tTipoIngreso", "nVarChar(1)  ", "lIncluidoImpuesto1", "Bit", _
                                          "lIncluidoImpuesto2", "Bit", "lIncluidoImpuesto3", "Bit", _
                                          "tCuentaContable", "nvarchar(15)  ", "tCentroCosto", "nvarchar(15) ", "tAreaOrigen", "nvarchar(3) ", "lAsumirIUE_IT", "Bit", "nRetencion1", "Float", _
                                          "nRetencion2", "Float", "nRetencion3", "Float", "nPRetencion1", "Float", _
                                          "nPRetencion2", "Float", "nPRetencion3", "Float", "lIncluidoRetencion1", "Bit", _
                                          "lIncluidoRetencion2", "Bit", "lIncluidoRetencion3", "Bit", "nICE", "Float", "nOtrosCargosInafecto", "Float")

    Isql = "sp_CargaDetalleDocumento '" & sDtemporal & "'"
    Set RsDtemporal = Lib.OpenRecordset(Isql, CnAlmacen)
    Set grdGrilla.DataSource = RsDtemporal
    
    LblItems = Format(RsDtemporal.RecordCount, "##0")
    
End Sub

'Public Sub PreparaTipoCambio()
'    'Consulta Tipo de Cambio
'    Isql = "SELECT [fFecha], [nCompra], [nVenta] FROM tTipoCambio " _
'           & " WHERE  fFecha='" & Format(DtpEmision.value, "MM/dd/yyyy") & "'"
'    Set RsTipoCambio = Lib.OpenRecordset(Isql, CnAlmacen)
'    If Not RsTipoCambio.EOF Then
'        RsTipoCambio.MoveLast
'        nTC = RsTipoCambio!nVenta
'    Else
'        MsgBox "El Tipo de Cambio de hoy" & vbCrLf & "aún no ha sido registrado...", vbInformation, sMensaje: Exit Sub
'        EstadoCabeceraCambio False
'        ActivaAdicion False
'    End If
'    LblTipoCambio.Caption = Format(nTC, "##0.000")
'End Sub

Sub EstadoCabecera(lValor As Boolean)
'    If lValor Then
'        LblPDoc.ForeColor = &H800000
'        LblPDocRef.ForeColor = &H800000
'    Else
'        LblPDoc.ForeColor = &H808080
'        LblPDocRef.ForeColor = &H808080
'    End If
    
    TxtSerie.Enabled = lValor
    TxtCorrelativo.Enabled = lValor
    'TxtSerieRef.Enabled = lValor
    'TxtCorrelativoRef.Enabled = lValor
    
    TxtGlosaDocumento.Enabled = lValor
    txtMontoControl.Enabled = lValor
    'TxtRUC.Enabled = lValor
    'TxtProveedor.Enabled = lValor
    'TxtNombre.Enabled = lValor
    
    DtpEmision.Enabled = lValor
    DtpIngreso.Enabled = lValor
    
    CboMoneda.Enabled = lValor
    CboArea.Enabled = lValor
    'CboReferencia.Enabled = lValor
    CboTipoIngreso.Enabled = lValor
    CboTipoDocumento.Enabled = lValor
    cboOperacion.Enabled = lValor  '07/02/2008
    cmdFind.Enabled = lValor
    txtTransporte.Enabled = lValor
    
    'txtCodigoControl.Enabled = lValor
    lblCodigoControl.Enabled = lValor
    'txtAutorizacion.Enabled = lValor
    lblCodigoAutori.Enabled = lValor
    'dtpFechCaducidad.Enabled = lValor
    
    dtcLibroCompra.Enabled = lValor
    
    dtpAnio.Enabled = lValor
    CmbMes.Enabled = lValor
    
    
End Sub

Sub EstadoCabeceraCambio(lValor As Boolean)
    CboTipoDocumento.Enabled = lValor
    'CboReferencia.Enabled = lValor
    CboMoneda.Enabled = lValor
    CboArea.Enabled = lValor
    CboTipoIngreso.Enabled = lValor
    
    TxtSerie.Enabled = lValor
    TxtCorrelativo.Enabled = lValor
    'TxtProveedor.Enabled = lValor
    'TxtNombre.Enabled = lValor
    'TxtSerieRef.Enabled = lValor
    'TxtCorrelativoRef.Enabled = lValor
    LblSumDescuento.Enabled = lValor
    TxtGlosaDocumento.Enabled = lValor
    dtcLibroCompra.Enabled = lValor
    txtMontoControl.Enabled = lValor
    'TxtRUC.Enabled = lValor
    
    LblPDoc.Enabled = lValor
    
    DtpIngreso.Enabled = lValor
    DtpEmision.Enabled = lValor
    
    cmdOpcion(7).Enabled = lValor
    cmdOpcion(4).Enabled = lValor
    cmdOpcion(0).Enabled = lValor
    
    If LblPDoc = "G" Or LblTotal = LblSaldo Then
        cmdOpcion(8).Enabled = False
    Else
        cmdOpcion(8).Enabled = lValor
    End If
    cmdFind.Enabled = lValor
    txtTransporte.Enabled = lValor
    CmdProcesar.Enabled = lValor
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    'cmdOpcion(0).Enabled = Activa  'Agregar
    cmdOpcion(1).Enabled = Activa
    'cmdOpcion(2).Enabled = Activa  'Anular
    
    'cmdOpcion(7).Enabled = True  'Canje
    
    'cmdOpcion(5).Enabled = Activa  'Emite
    
'    If LblPDoc = "G" Or LblTotal = LblSaldo Then
'        cmdOpcion(8).Enabled = False  'Historial
'    Else
'        cmdOpcion(8).Enabled = Activa  'Historial
'    End If
    cmdPago.Enabled = Activa
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Liberar
    sUsuarioAutoriza = ""
    Set frmDocumentoCompra = Nothing
    'sForm = ""
End Sub

Sub Activa(valor As Boolean)
    cmdOpcion(1).Enabled = valor
End Sub

Sub ActivaAdicion(ByVal valor As Boolean)
    CmdAdicion(0).Enabled = valor
    'CmdAdicion(1).Enabled = Valor
    CmdAdicion(2).Enabled = valor
    If Sw Then
       CmdAdicion(3).Enabled = True
    Else
       CmdAdicion(3).Enabled = False
    End If
    'cmdDescuento.Enabled = Valor
    'Bolivia
    Select Case pais
        Case "000", "004", "005": cmdOpcion(10).Enabled = valor
        Case "001": cmdOpcion(10).Enabled = False
        Case "002": cmdOpcion(10).Enabled = valor
    End Select
End Sub




Public Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If SwTmp Then
'        CalculaTmp
'    Else
'        If Not lAdvertenciaDifImpuestos Then
'            CalculaTabla
'        End If
'    End If
End Sub

Sub CalculaTabla()
    Select Case pais
            Case "000", "002", "003", "004", "005"
                If Not lNavegar And Not lLoad Then
                    Isql = "Select Sum(nImpuesto1) as nImpuesto1, Sum(nImpuesto2) as nImpuesto2, Sum(nImpuesto3) as nImpuesto3, Sum(nNeto) as nNeto, Sum(nDescuento+ nDescuento*nPorcentaje1/100 + nDescuento*nPorcentaje2/100+ nDescuento*nPorcentaje3/100) as nDescuento From DDocumento Where tCorrelativo='" & LblCorrelativo & "'"
'                    Isql = "Select Sum(nImpuesto1) as nImpuesto1, Sum(nImpuesto2) as nImpuesto2, Sum(nImpuesto3) as nImpuesto3, Sum(nNeto) as nNeto, Sum(nDescuento) as nDescuento From DDocumento Where tCorrelativo='" & LblCorrelativo & "'"
                    Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
'                    If Rs.RecordCount > 0 Then
'                        wImp1 = IIf(IsNull(Rs!nImpuesto1), 0, Rs!nImpuesto1)
'                        wImp2 = IIf(IsNull(Rs!nImpuesto2), 0, Rs!nImpuesto2)
'                        wImp3 = IIf(IsNull(Rs!nImpuesto3), 0, Rs!nImpuesto3)
'                        wSubTotal = IIf(IsNull(Rs!nNeto), 0, Rs!nNeto)
'                        wDesc = IIf(IsNull(Rs!nDescuento), 0, Rs!nDescuento)
'                    Else
'                        wImp1 = 0
'                        wImp2 = 0
'                        wImp3 = 0
'                        wSubTotal = 0
'                        wDesc = 0
'                    End If
                    
                    If lActualizaImpuesto Then
'                        If lAdvertenciaDifImpuestos Then
'                            If (LblTotImp1 <> Format(wImp1, "##,##0.00") Or LblTotImp2 <> Format(wImp2, "##,##0.00") Or LblTotImp3 <> Format(wImp3, "##,##0.00")) Then
'                                If MsgBox("Existen diferencias entre los impuestos de la cabecera y el detalle." & vbCrLf & "¿Desea actualizar los valores de los impuestos en la cabecera del documento " & LblPDoc.Caption & TxtSerie & TxtCorrelativo & "?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
'                                    wImp1 = LblTotImp1
'                                    wImp2 = LblTotImp2
'                                    wImp3 = LblTotImp3
'                                Else
'                                    LblTotImp1 = Format(wImp1, "##,##0.00")
'                                    LblTotImp2 = Format(wImp2, "##,##0.00")
'                                    LblTotImp3 = Format(wImp3, "##,##0.00")
'                                End If
'                                Screen.MousePointer = vbDefault
'                            Else
'                                LblTotImp1 = Format(wImp1, "##,##0.00")
'                                LblTotImp2 = Format(wImp2, "##,##0.00")
'                                LblTotImp3 = Format(wImp3, "##,##0.00")
'                            End If
'                        Else
'                            LblTotImp1 = Format(wImp1, "##,##0.00")
'                            LblTotImp2 = Format(wImp2, "##,##0.00")
'                            LblTotImp3 = Format(wImp3, "##,##0.00")
'                        End If
                    Else
'                        LblTotImp1 = Format(wImp1, "##,##0.00")
'                        LblTotImp2 = Format(wImp2, "##,##0.00")
'                        LblTotImp3 = Format(wImp3, "##,##0.00")
                    End If
                    
'                    LblSubTot = Format(wSubTotal, "##,##0.00")
'                    LblTotal = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(txtTransporte.Text), "##,##0.00")
'                    LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(lblTotLey1) + CDbl(txtTransporte.Text), "##,##0.00")
'
'                    LblSumDescuento = Format(wDesc, "#,##0.00")

                End If
            
    End Select
'    If lActualiza Then
'        LblSaldo = LblTotal
'        lActualiza = False
'    End If
End Sub




Private Sub LblICE_Change()
        If Not LblICE.Caption = "" And wICE = True Then
                Isql = "exec SP_AgregarICE_Tabla " & vICE & "," & val(LblICE.Caption) & ", '" & LblCorrelativo & "'"
                CnAlmacen.Execute Isql
                vICE = val(LblICE.Caption)
                RsDocumentoD.Requery
                lLoad = False
                CalculaTabla
                Actualiza
                lLoad = True
            wICE = False
        End If
End Sub

Private Sub LblICE_DblClick()
    If wVaRegCompra Then
        If wAfectoICE Then
            If LblEstado.Caption = "GENERADO" Or Sw Then
                wICE = True
                TxtICE.Visible = True
                TxtICE.Text = LblICE.Caption
                TxtICE.SetFocus
                TxtICE.SelStart = 0
                TxtICE.SelLength = Len(TxtICE)
            End If
        End If
    End If
End Sub


Private Sub LblPDoc_Change()
    Isql = "Select tCodigoTipoDocumento,tDetallado,tResumido,lRegistroCompras FROM tTipoDocumento " _
         & " Where tResumido='" & LblPDoc.Caption & "'"
    Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
    If Rs.EOF Then Exit Sub
'    wVaRegCompra = Rs!lRegistroCompras
'    If wVaRegCompra Then
'        txtCodigoControl.backColor = &HF0EADB
'        txtAutorizacion.backColor = &HF0EADB
'    Else
'        txtCodigoControl.backColor = &HFFFFFF
'        txtAutorizacion.backColor = &HFFFFFF
'    End If
    CboTipoDocumento.BoundText = Rs!tCodigoTipoDocumento
    If TxtSerie.Visible = True And TxtSerie.Enabled = True Then TxtSerie.SetFocus
End Sub



Private Sub lblRedondeo_Click()
    If LblEstado.Caption = "---" Then Exit Sub
    'If cmdOpcion(3).Visible = True And cmdOpcion(3).Enabled = True Then Exit Sub
    iTBox = 6
    
End Sub

Private Sub LblSubTot_Click()
'    If Not Sw Then
'        If (frmDocumento.RsCabecera!CabEstadoDocumento = "01" Or frmDocumento.RsCabecera!CabEstadoDocumento = "02") And Trim(LblCorrelativo) <> "" And CDbl(LblTotal) > 0 Then
'            frmMontoDesglose.Show vbModal
'        End If
'    End If
End Sub

Private Sub LblTotal_Change()
    '&H80000005&
    If txtMontoControl = LblTotal Then
        txtMontoControl.backColor = &HC0FFC0
    Else
        txtMontoControl.backColor = &HC0C0FF
    End If
End Sub

Private Sub LblTotal_Click()
    If LblEstado.Caption = "---" Then Exit Sub
    'If cmdOpcion(3).Visible = True And cmdOpcion(3).Enabled = True Then Exit Sub
    iTBox = 5
'    If Not Sw Then
'        If (frmDocumento.RsCabecera!CabEstadoDocumento = "01" Or frmDocumento.RsCabecera!CabEstadoDocumento = "02") And Trim(LblCorrelativo) <> "" And CDbl(LblTotal) > 0 Then
'            frmAjusteTotal.Show vbModal
'            AjusteTotal
'        End If
'    End If
End Sub

Private Sub LblTotImp1_Click()
    If LblEstado.Caption = "---" Then Exit Sub
    'If cmdOpcion(3).Visible = True And cmdOpcion(3).Enabled = True Then Exit Sub
    iTBox = 2
'    If Not Sw Then
'        If (frmDocumento.RsCabecera!CabEstadoDocumento = "01" Or frmDocumento.RsCabecera!CabEstadoDocumento = "02") And Trim(LblCorrelativo) <> "" And CDbl(LblTotal) > 0 Then
'            frmAjusteTotal.Show vbModal
'            AjusteTotal
'        End If
'    End If
End Sub

Private Sub LblTotImp2_Click()
    If LblEstado.Caption = "---" Then Exit Sub
    'If cmdOpcion(3).Visible = True And cmdOpcion(3).Enabled = True Then Exit Sub
    iTBox = 3
'    If Not Sw Then
'        If (frmDocumento.RsCabecera!CabEstadoDocumento = "01" Or frmDocumento.RsCabecera!CabEstadoDocumento = "02") And Trim(LblCorrelativo) <> "" And CDbl(LblTotal) > 0 Then
'            frmAjusteTotal.Show vbModal
'            AjusteTotal
'        End If
'    End If
End Sub

Private Sub LblTotImp3_Click()
    If LblEstado.Caption = "---" Then Exit Sub
    'If cmdOpcion(3).Visible = True And cmdOpcion(3).Enabled = True Then Exit Sub
    iTBox = 4
'    If Not Sw Then
'        If (frmDocumento.RsCabecera!CabEstadoDocumento = "01" Or frmDocumento.RsCabecera!CabEstadoDocumento = "02") And Trim(LblCorrelativo) <> "" And CDbl(LblTotal) > 0 Then
'            frmAjusteTotal.Show vbModal
'            AjusteTotal
'        End If
'    End If
End Sub



Private Sub txtAutorizacion_LostFocus()
    If pais = "001" Then
        GeneraMensaje
    End If
End Sub



Private Sub txtCorrelativo_Change()
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xTxtCorrelativo = True
'    End If
End Sub

Private Sub TxtCorrelativo_GotFocus()
    TxtCorrelativo.SelStart = 0
    TxtCorrelativo.SelLength = Len(TxtCorrelativo)
End Sub

Private Sub TxtCorrelativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'txtRuc.SetFocus
        DtpEmision.SetFocus
    Else
        If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122) And KeyAscii <> 8 Then
'        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            MsgBox "Ingrese únicamente números y letras...", vbInformation, sMensaje
            KeyAscii = 0
            TxtCorrelativo.SetFocus
        End If
    End If
End Sub

Private Sub txtCorrelativo_LostFocus()
    If pais <> "002" Then
        If Len(Me.TxtCorrelativo) <> 0 Then
            TxtCorrelativo = Mid("000000000", 1, 9 - Len(Trim(TxtCorrelativo))) + Trim(TxtCorrelativo)
        Else
            TxtCorrelativo = ""
        End If
    Else
        If Len(Me.TxtCorrelativo) <> 0 Then
            If Len(Trim(TxtCorrelativo)) > 8 Then
                TxtCorrelativo.Text = Left(TxtCorrelativo.Text, 8)
            End If
            TxtCorrelativo = Mid("00000000", 1, 8 - Len(Trim(TxtCorrelativo))) + Trim(TxtCorrelativo)
        Else
            TxtCorrelativo = ""
        End If
    End If
    'If cmdOpcion(3).Visible Then Exit Sub
End Sub

Sub GeneraMensaje()
    If swMover Then Exit Sub
    Dim RsCorrela As Recordset
    
    'No hacer el select con el correlativo, se busca solamente documentos iguales...
    If pais <> "001" Then
        If Sw Then
            Isql = "Select * From mDocumento Where tDocumento='" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' And tCodigoProveedor='" & TxtProveedor.Text & "' Order by tCorrelativo"
        Else
            Isql = "Select * From mDocumento Where tDocumento='" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' And tCodigoProveedor='" & TxtProveedor.Text & "' and tDocumento not in ('" & sDocumento & "') Order by tCorrelativo"
        End If
    Else
        Isql = "Select * From mDocumento Where tDocumento='" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' And tCodigoProveedor='" & TxtProveedor.Text & "' and tCodigoAutorizacion='' Order by tCorrelativo"
    End If
    Set RsCorrela = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsCorrela.EOF Then
        RsCorrela.MoveLast
        If RsCorrela!tEstadoDocumento <> "04" Then
            MsgBox "El documento digitado ya está registrado. No Procede...", vbCritical, sMensaje
            TxtProveedor = ""
            TxtNombre = ""
            TxtRUC = ""
            'TxtDiasCredito = ""
            If TxtProveedor.Enabled Then TxtProveedor.SetFocus
            Exit Sub
        End If
    End If
End Sub

Function VerificarExistenciaDocumento() As Boolean
    If swMover Then Exit Function
    Dim RsCorrela As Recordset
    
    'No hacer el select con el correlativo, se busca solamente documentos iguales...
    If pais <> "001" Then
        Isql = "Select * From mDocumento Where tDocumento='" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' And tCodigoProveedor='" & TxtProveedor & "' and tcorrelativo<>'" & LblCorrelativo & "'  Order by tCorrelativo"
    Else
        Isql = "Select * From mDocumento Where tDocumento='" & Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo) & "' And tCodigoProveedor='" & TxtProveedor & "' and tcorrelativo<>'" & LblCorrelativo & "'  and tCodigoAutorizacion='' Order by tCorrelativo"
    End If
    Set RsCorrela = Lib.OpenRecordset(Isql, CnAlmacen)
    If Not RsCorrela.EOF Then
        RsCorrela.MoveLast
        If RsCorrela!tEstadoDocumento <> "04" Then
            MsgBox "El documento digitado ya está registrado. No Procede...", vbCritical, sMensaje
'            TxtProveedor = ""
'            TxtNombre = ""
'            txtRuc = ""
'            TxtDiasCredito = ""
            CboTipoDocumento.BoundText = ""
            LblPDoc.Caption = ""
            If TxtProveedor.Enabled Then TxtProveedor.SetFocus
            VerificarExistenciaDocumento = False
            Exit Function
        Else
            VerificarExistenciaDocumento = True
        End If
    Else
        VerificarExistenciaDocumento = True
    End If
End Function





Private Sub txtDobserva_Change()
    If Len(txtDobserva) > txtDobserva.MaxLength Then
        MsgBox "Se ha llegado al límite de caracteres...", vbInformation, sMensaje
        Exit Sub
    End If
End Sub

Private Sub TxtGlosaDocumento_Change()
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xTxtGlosaDocumento = True
'    End If
End Sub

Private Sub TxtGlosaDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdAdicion(0).SetFocus
End Sub

Private Sub txtICE_KeyPress(KeyAscii As Integer)
'    swPunto = ChrBuscaPunto(txtICE.Text)
'    If Trim(txtICE) = "" Then swPunto = False
'    If KeyAscii = 13 Then
'        If txtICE.Text = "" Then txtICE.Text = "0"
'        LblICE.Caption = Format(txtICE.Text, "##,##0.00"):    txtICE.Visible = False
'    Else
'        If KeyAscii <> 27 Then
'            PictureNumero KeyAscii
'       End If
'    End If
'    If KeyAscii = 27 Then LblICE.Caption = LblICE.Caption: txtICE.Visible = False
End Sub

Private Sub txtMontoControl_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then TxtGlosaDocumento.SetFocus
'
'    swPunto = ChrBuscaPunto(txtMontoControl.Text)  'If Trim(TxtCantidad) = "" Then swPunto = False
'    If Trim(txtMontoControl) = "" Then swPunto = False
'    If KeyAscii = 13 Then
'       TxtGlosaDocumento.SetFocus
'    Else
'       PictureNumero KeyAscii      'Para evaluar si lo digitado es un numero
'    End If
    
End Sub

'Private Sub txtMontoControl_Change()
'    '&H80000005&
'
'End Sub

Private Sub txtMontoControl_LostFocus()
    txtMontoControl = Format(txtMontoControl, "##,##0.00")
    If txtMontoControl = LblTotal Then
        txtMontoControl.backColor = &HC0FFC0
    Else
        txtMontoControl.backColor = &HC0C0FF
    End If
End Sub

Private Sub TxtNombre_Change()
'    If Len(Trim(TxtNombre)) = 1 Then
'        'xISQL = ""
'        Isql = "Select tCodigoProveedor As Codigo,tRazonSocial As Descripcion,tIndicadorTributario, isnull(lAfectoICE,0) as lAfectoICE From tProveedor  where lActivo=1  Order By tCodigoProveedor"
'        'sBuscar = "Compania"
'        'frmBusca.lEntra = True
'        frmBusca.txtCriterio = Trim(TxtNombre)
'        frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
'        lCriterio = False
'        frmBusca.Show vbModal
'        If Sw Then
'            TxtDiasCredito = Calcular("Select nCredito As Codigo From tProveedor Where tCodigoProveedor='" & TxtProveedor & "'", CnAlmacen)
'        Else
'            TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
'        End If
'    End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DtpIngreso.SetFocus 'TxtProveedor.SetFocus
End Sub

Private Sub txtNombre_LostFocus()
    'If cmdOpcion(3).Visible Then Exit Sub
    GeneraMensaje
End Sub


Private Sub txtRuc_Change()
'    If Len(Trim(TxtRUC)) = 1 Then
'        'xISQL = ""
'        Isql = "Select tCodigoProveedor As Codigo,tRazonSocial As Descripcion,tIndicadorTributario, isnull(lAfectoICE,0) as lAfectoICE From tProveedor where lActivo=1 Order By tCodigoProveedor"
'        'sBuscar = "Tributario"
'        'frmBusca.lEntra = True
'        frmBusca.txtCriterio = Trim(TxtRUC)
'        frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
'        lCriterio = False
'        frmBusca.Show vbModal
'        If Sw Then
'            TxtDiasCredito = Calcular("Select nCredito As Codigo From tProveedor Where tCodigoProveedor='" & TxtProveedor & "'", CnAlmacen)
'        Else
'            TxtDiasCredito = Format(DtpFPrograma - DtpIngreso, "##0")
'        End If
'    End If
   
    'If cmdOpcion(3).Visible Then Exit Sub
    GeneraMensaje
End Sub

Private Sub txtProveedor_Change()
    If Len(Trim(TxtProveedor)) Then
        frmBusca.lEntra = True
    Else
        Exit Sub
    End If
    
    If frmBusca.lEntra Then
        If Len(Trim(TxtProveedor)) = 1 Then
            Isql = "Select tCodigoProveedor As Codigo,tRazonSocial As Descripcion,tIndicadorTributario From tProveedor where lActivo=1 Order By tCodigoProveedor"
            sBuscar = "Proveedor"
            frmBusca.txtCriterio = Trim(TxtProveedor)
            frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
                  
            lCriterio = False
            frmBusca.Show vbModal
            
        End If
    End If
    
    
End Sub

Private Sub TxtProveedor_LostFocus()
    'If cmdOpcion(3).Visible Then Exit Sub
    GeneraMensaje
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtNombre.SetFocus
End Sub

Private Sub txtRuc_LostFocus()
    'If cmdOpcion(3).Visible Then Exit Sub
    GeneraMensaje
End Sub

Private Sub txtSerie_Change()
'    If cmdOpcion(3).Visible Then
'        cmdOpcion(4).Enabled = True
'        'xTxtSerie = True
'    End If
End Sub

Private Sub TxtSerie_GotFocus()
    TxtSerie.SelStart = 0
    TxtSerie.SelLength = Len(TxtSerie)
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCorrelativo.SetFocus
    Else
        If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122) And KeyAscii <> 8 Then
            MsgBox "Ingrese únicamente números y letras...", vbInformation, sMensaje
            KeyAscii = 0
            TxtSerie.SetFocus
        End If
    End If
End Sub

Private Sub txtSerie_LostFocus()
    If pais <> "002" Then
'       If Trim(txtSerie) = "" Or Not IsNumeric(txtSerie) Then txtSerie = 0
        If Trim(TxtSerie) = "" Then TxtSerie = 0
'        If CDbl(TxtSerie) < 0 Then TxtSerie = 0
'        If Len(Trim(str(val(Me.TxtSerie)))) > 5 Then TxtSerie.Text = Left(Me.TxtSerie.Text, 5)
        If Len(Trim(Me.TxtSerie)) > 5 Then TxtSerie.Text = Left(Me.TxtSerie.Text, 5)
        If Trim(LblPDoc) <> "" Then
'           TxtSerie = Mid("00000", 1, 5 - Len(Trim(str(val(Me.TxtSerie))))) + Trim(str(val(Me.TxtSerie)))
           TxtSerie = Mid("00000", 1, 5 - Len(Trim(Me.TxtSerie))) + Trim(Me.TxtSerie)
        Else
           TxtSerie = ""
        End If
    Else
        If Trim(TxtSerie) = "" Or Not IsNumeric(TxtSerie) Then TxtSerie = 0
        If CDbl(TxtSerie) < 0 Then TxtSerie = 0
        If Trim(LblPDoc) <> "" Then
           'TxtSerie = Mid("000000", 1, 6 - Len(Trim(str(val(Me.TxtSerie))))) + Trim(str(val(Me.TxtSerie)))
           TxtSerie = Mid("000000", 1, 6 - Len(Trim(Me.TxtSerie))) + Trim(Me.TxtSerie)
        Else
           TxtSerie = ""
        End If
    End If
End Sub

Private Sub TxtSerieRef_Change()
    If cmdOpcion(3).Visible Then
        cmdOpcion(4).Enabled = True
        'xTxtSerieRef = True
    End If
End Sub







Private Sub PrevioAnula()
    If CDbl(LblTotal) > CDbl(LblSaldo) Then
        MsgBox "No es posible desprocesar el documento. Existe pagos rea-" & vbCrLf & _
               "lizados en Cuenta Corriente o Notas de Crédito aplicadas...", vbCritical, sMensaje
        wEnter = False
        Exit Sub
    End If
    If CDbl(LblSaldo) = 0 Then
        MsgBox "No es posible desprocesar el documento. El" & vbCrLf & "documento ya se encuentra cancelado...", vbCritical, sMensaje
        wEnter = False
        Exit Sub
    End If
    
    Dim sUltimo     As String
    Dim sSinStock   As String
    sSinStock = "No es posible desprocesar el documento." & vbCrLf & "No hay stock suficiente para descargar:" & vbCrLf & vbCrLf
    
   'Sólo se verifican los documentos que sean los que generan movimiento en el Kardex...
   'Los documentos que no lo sean, no se verifican y se anulan directamente...
   
    'If Left(CboTipoIngreso, 1) = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
   
       If CboArea.BoundText = "000" Then       'Almacen Central
     
          '1) Verifica si es el último documento ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

           sUltimo = Calcular("Select Max(tCorrelativo) As Codigo From mDocumento", CnAlmacen)
           If sUltimo = LblCorrelativo Then
              
             '3) Verifica que cada item a revertir tenga stock ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                 
                 wEnter = True
                 If Sw Then
                    RsDtemporal.MoveFirst
                    Do While Not RsDtemporal.EOF
                       sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
                       
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                             
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDtemporal.MoveNext
                    Loop
                 Else
                    RsDocumentoD.MoveFirst
                    Do While Not RsDocumentoD.EOF
                       sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                            
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDocumentoD.MoveNext
                    Loop
                 End If
                 
                 grdGrilla.MoveFirst
                 If Not wEnter Then
                    MsgBox sSinStock & vbCrLf & "Inténtelo cuando haya stock...", vbCritical, sMensaje
                    wEnter = False
                    Exit Sub
                 End If
                 
           Else
           
             '3) Verifica que cada item a revertir tenga stock ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           
                 wEnter = True
                 If Sw Then
                    RsDtemporal.MoveFirst
                    Do While Not RsDtemporal.EOF
                       sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                           
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDtemporal.MoveNext
                    Loop
                 Else
                    RsDocumentoD.MoveFirst
                    Do While Not RsDocumentoD.EOF
                       sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                            
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDocumentoD.MoveNext
                    Loop
                 End If
                 
                 grdGrilla.MoveFirst
                 If Not wEnter Then
                    MsgBox sSinStock & vbCrLf & "Inténtelo cuando haya stock...", vbCritical, sMensaje
                    wEnter = False
                    Exit Sub
                 End If
           End If
       
       Else         '~ Sub-Almacenes ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       
          '1) Verifica si es el último documento ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           
           sUltimo = Calcular("Select Max(tCorrelativo) As Codigo From mDocumento", CnAlmacen)
           If sUltimo = LblCorrelativo Then
              
             '3) Verifica que cada item a revertir tenga stock ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                 
                 wEnter = True
                 If Sw Then
                    RsDtemporal.MoveFirst
                    Do While Not RsDtemporal.EOF
                       sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
                          sArti = Calcular("Select tDetallado As codigo From tProducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                          
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tSubStock Where tCodigoProducto='" & sCodi & "' And tCodigoSubArea='" & CboArea.BoundText & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDtemporal.MoveNext
                    Loop
                 Else
                    RsDocumentoD.MoveFirst
                    Do While Not RsDocumentoD.EOF
                       sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                            
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tSubStock Where tCodigoProducto='" & sCodi & "' And tCodigoSubArea='" & CboArea.BoundText & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDocumentoD.MoveNext
                    Loop
                 End If
                 
                 grdGrilla.MoveFirst
                 If Not wEnter Then
                    MsgBox sSinStock & vbCrLf & "Inténtelo cuando haya stock...", vbCritical, sMensaje
                    wEnter = False
                    Exit Sub
                 End If
           
           Else
           
             '3) Verifica que cada item a revertir tenga stock ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                 
                 wEnter = True
                 If Sw Then
                    RsDtemporal.MoveFirst
                    Do While Not RsDtemporal.EOF
                       sCodi = IIf(Not IsNull(RsDtemporal!tCodigoProducto), RsDtemporal!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDtemporal!nCantidad), RsDtemporal!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                            
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tSubStock Where tCodigoProducto='" & sCodi & "' And tCodigoSubArea='" & CboArea.BoundText & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If           'NUEVO END IF -> 4-Mar-2005
                       
                       RsDtemporal.MoveNext
                    Loop
                 Else
                    RsDocumentoD.MoveFirst
                    Do While Not RsDocumentoD.EOF
                       sCodi = IIf(Not IsNull(RsDocumentoD!tCodigoProducto), RsDocumentoD!tCodigoProducto, "")
                      
                      'Nuevo IF -> 4-Mar-2005
                       'If Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'") = Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & Left(CboTipoIngreso, 1) & "'", CnAlmacen) Then
                       sTipoProd = Calcular("SELECT TTIPOPRODUCTO AS CODIGO FROM TPRODUCTO WHERE TCODIGOPRODUCTO='" & sCodi & "'", CnAlmacen)
                       If Calcular("Select tCodigo As Codigo From tTabla Where tTabla='TIPOINGRESO' And tvalor='1' And tCodigo='" & sTipoProd & "'", CnAlmacen) = sTipoProd Then
                       
                          sCant = IIf(Not IsNull(RsDocumentoD!nCantidad), RsDocumentoD!nCantidad, "")
                          sArti = Calcular("Select tDetallado as codigo from tproducto where tcodigoproducto='" & sCodi & "'", CnAlmacen)
                            
                          dStockArt = CDbl(Calcular("Select nStockActual As Codigo From tSubStock Where tCodigoProducto='" & sCodi & "' And tCodigoSubArea='" & CboArea.BoundText & "'", CnAlmacen))
                          dMermaArt = CDbl(Calcular("Select nMerma As Codigo From tProducto Where tCodigoProducto='" & sCodi & "'", CnAlmacen))
                          If dStockArt + (CDbl(sCant) * (dMermaArt / 100)) < CDbl(sCant) Then
                             sSinStock = sSinStock & "- " & sArti & vbCrLf
                             wEnter = False
                          End If
                       
                       End If       'NUEVO END IF -> 4-Mar-2005
                       
                       RsDocumentoD.MoveNext
                    Loop
                 End If
                 
                 grdGrilla.MoveFirst
                 If Not wEnter Then
                    MsgBox sSinStock & vbCrLf & "Inténtelo cuando haya stock...", vbCritical, sMensaje
                    wEnter = False
                    Exit Sub
                 End If
           
           End If
       
       End If
    
    'Else
    
    '   wEnter = False
    
    'End If
End Sub

Sub ValidaCierre()
    sCerrado = Calcular("Select lCierre As Codigo From tCierre Where tAno='" & Trim(str(Year(DtpIngreso))) & "' And tMes='" & Trim(str(Month(DtpIngreso))) & "'", CnAlmacen)
    If sCerrado <> "0" Then
        'Está en la tabla de cierre...
        If sCerrado = "Verdadero" Then
            Screen.MousePointer = vbDefault
            MsgBox "No es posible realizar ninguna operación con" & vbCrLf & "este documento por encontrarse el mes cerrado...", vbCritical, sMensaje
            Exit Sub
        End If
    End If
End Sub

Sub ValidaPlazo()
    Dim iPlazoIngDoc As Integer
    
    'swPasar = True
    iPlazoIngDoc = Calcular("Select nPlazoIngDoc As Codigo From tParametro", CnAlmacen)
    If Date - DtpEmision > iPlazoIngDoc Then
        Screen.MousePointer = vbDefault
        MsgBox "No está permitido ingresar un documento pasados" & vbCrLf & "los " & Trim(str(iPlazoIngDoc)) & " días desde su fecha de emisión. Verifique...", vbCritical, sMensaje
        If DtpEmision.Enabled Then DtpEmision.SetFocus
        'swPasar = False
    End If
End Sub

'Sub ValidaPlazoRegistro()
'    Dim iPlazoIngDoc As Integer
'
'    swPasar = True
'    iPlazoIngDoc = Calcular("Select nPlazoRegDoc As Codigo From tParametro", CnAlmacen)
'    If date - DtpEmision > iPlazoIngDoc Then
'        Screen.MousePointer = vbDefault
'        MsgBox "No está permitido registrar un documento pasados" & vbCrLf & "los " & Trim(str(iPlazoIngDoc)) & " días desde su fecha de emisión. Verifique...", vbCritical, sMensaje
'        If DtpEmision.Enabled Then DtpEmision.SetFocus
'        swPasar = False
'    End If
'End Sub

Function ValidaDuplicidad() As Boolean
    Dim tEstado As String
    tEstado = Calcular("Select tEstadoDocumento as Codigo FROM MDOCUMENTO where tCorrelativo='" & LblCorrelativo & "'", CnAlmacen)
    ValidaDuplicidad = IIf(tEstado = "01", False, True)
End Function

Function ValidaDuplicidadDes() As Boolean
    Dim tEstado As String
    tEstado = Calcular("Select tEstadoDocumento as Codigo FROM MDOCUMENTO where tCorrelativo='" & LblCorrelativo & "'", CnAlmacen)
    ValidaDuplicidadDes = IIf(tEstado = "02", False, True)
End Function


Public Sub ActDetalle()
      'Elimina primero el detalle que pueda existir para ser (nuevamente) insertado...
       CnAlmacen.Execute "Delete From dDocumento Where tCorrelativo='" & LblCorrelativo & "'"
       
       sitem = 0
       If RsDocumentoD.RecordCount > 0 Then
          RsDocumentoD.MoveFirst
       End If
       Set mdDocumento = New ClsDocumento
       Do While Not RsDocumentoD.EOF

          sitem = sitem + 1

          mdDocumento.InsdDocumento Trim(LblPDoc) & Trim(TxtSerie) & Trim(TxtCorrelativo), Trim(TxtProveedor.Text), sitem, IIf(IsNull(RsDocumentoD!tCodigoProducto), "", RsDocumentoD!tCodigoProducto), IIf(IsNull(RsDocumentoD!tUnidadEntrada), "", RsDocumentoD!tUnidadEntrada), IIf(IsNull(RsDocumentoD!nCantidad), 0, RsDocumentoD!nCantidad), IIf(IsNull(RsDocumentoD!nPrecio), 0, RsDocumentoD!nPrecio), IIf(IsNull(RsDocumentoD!nPVenta), 0, RsDocumentoD!nPVenta), IIf(IsNull(RsDocumentoD!nNeto), 0, RsDocumentoD!nNeto), IIf(IsNull(RsDocumentoD!nDescuento), 0, RsDocumentoD!nDescuento), IIf(IsNull(RsDocumentoD!nRecargo), 0, RsDocumentoD!nRecargo), _
                                    IIf(IsNull(RsDocumentoD!nImpuesto1), 0, RsDocumentoD!nImpuesto1), IIf(IsNull(RsDocumentoD!nImpuesto2), 0, RsDocumentoD!nImpuesto2), IIf(IsNull(RsDocumentoD!nImpuesto3), 0, RsDocumentoD!nImpuesto3), IIf(IIf(IsNull(RsDocumentoD!lIncluidoImpuesto1), False, RsDocumentoD!lIncluidoImpuesto1) = True, RsDocumentoD!nPorcentaje1, 0), IIf(IIf(IsNull(RsDocumentoD!lIncluidoImpuesto2), False, RsDocumentoD!lIncluidoImpuesto2) = True, RsDocumentoD!nPorcentaje2, 0), IIf(IIf(IsNull(RsDocumentoD!lIncluidoImpuesto3), False, RsDocumentoD!lIncluidoImpuesto3) = True, RsDocumentoD!nPorcentaje3, 0), _
                                    IIf(IsNull(RsDocumentoD!nTotal), 0, RsDocumentoD!nTotal), "01", IIf(IsNull(CboTipoIngreso.BoundText), "", CboTipoIngreso.BoundText), IIf(IsNull(RsDocumentoD!lIncluidoImpuesto1), False, RsDocumentoD!lIncluidoImpuesto1), IIf(IsNull(RsDocumentoD!lIncluidoImpuesto2), False, RsDocumentoD!lIncluidoImpuesto2), IIf(IsNull(RsDocumentoD!lIncluidoImpuesto3), False, RsDocumentoD!lIncluidoImpuesto3), LblCorrelativo.Caption, IIf(IsNull(RsDocumentoD!TCUENTACONTABLE), "", RsDocumentoD!TCUENTACONTABLE), IIf(IsNull(RsDocumentoD!tCentroCosto), "", RsDocumentoD!tCentroCosto), IIf(IsNull(RsDocumentoD!tAreaOrigen), "", RsDocumentoD!tAreaOrigen), "I", _
                                    IIf(IsNull(RsDocumentoD!lAsumirIUE_IT), 0, RsDocumentoD!lAsumirIUE_IT), IIf(IsNull(RsDocumentoD!nRetencion1), "", RsDocumentoD!nRetencion1), IIf(IsNull(RsDocumentoD!nRetencion2), "", RsDocumentoD!nRetencion2), IIf(IsNull(RsDocumentoD!nRetencion3), "", RsDocumentoD!nRetencion3), IIf(IsNull(RsDocumentoD!nPRetencion1), "", RsDocumentoD!nPRetencion1), IIf(IsNull(RsDocumentoD!nPRetencion2), "", RsDocumentoD!nPRetencion2), IIf(IsNull(RsDocumentoD!nPRetencion3), "", RsDocumentoD!nPRetencion3), IIf(IsNull(RsDocumentoD!lIncluidoRetencion1), "", RsDocumentoD!lIncluidoRetencion1), IIf(IsNull(RsDocumentoD!lIncluidoRetencion2), "", RsDocumentoD!lIncluidoRetencion2), IIf(IsNull(RsDocumentoD!lIncluidoRetencion3), "", RsDocumentoD!lIncluidoRetencion3), _
                                    IIf(IsNull(RsDocumentoD!nICE), 0, RsDocumentoD!nICE), IIf(IsNull(RsDocumentoD!nOtrosCargosInafecto), 0, RsDocumentoD!nOtrosCargosInafecto)

          RsDocumentoD.MoveNext
          
       Loop
       Set mdDocumento = Nothing
       
       'Guarda la percepcion en el detalle del documento - ddocumento
       'cambio jesus 2014-09-11
       If nLey1 <> 0 Then
            Isql = "Exec spSavePerceptionDetails '" & LblCorrelativo.Caption & "'"
            CnAlmacen.Execute Isql
        End If
       'fin cambio
       
       'JESUS
       'If FrmDocumento.RsCabecera!CabEstadoDocumento = "01" Then
            If nLeyAD <> 0 Then
                 Isql = "Exec spSaveLeyADDetails '" & LblCorrelativo.Caption & "'"
                 CnAlmacen.Execute Isql
            End If
       'End If
       '------------02.11.2017-----------------
       
       lNavegar = False
       lLoad = False
       CalculaTabla
       'LblSaldo = Format(wSubTotal + wImp1 + wImp2 + wImp3 + CDbl(lblRedondeo) + CDbl(txtTransporte.Text), "##,##0.00")
       Actualiza
             
'       frmDocumento.RsCabecera.Requery
'       If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.MoveFirst
'       frmDocumento.RsCabecera.Find "tCorrelativo='" & LblCorrelativo & "'"
'       If frmDocumento.RsCabecera.EOF Then frmDocumento.RsCabecera.MoveLast
End Sub





Sub ActivarNavegar(Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
End Sub

Function VerificaAbonosNC(Prov As String, Docu As String) As Double
    Dim dAbonos, dNotasC As Double
    'dAbonos = CDbl(Calcular("SELECT tCodigoProveedor as Codigo,tDocumento,'nTotal'=CASE WHEN TMONEDA='01' THEN NIMPORTE ELSE NIMPORTE*NCAMBIO END From MPAGO Where tCodigoProveedor='" & Prov & "' And tDocumento='" & Docu & "'", CnAlmacen))
    'dNotasC = CDbl(Calcular("SELECT tCodigoProveedor as Codigo,tDocumento,'nTotal'=CASE WHEN TMONEDA='01' THEN NTOTAL ELSE NTOTAL*NTIPOCAMBIO END From MNOTACREDITO Where tCodigoProveedor='" & Prov & "' And tDocumento='" & Docu & "'", CnAlmacen))
    'VerificaAbonosNC = dAbonos + dNotasC
    dAbonos = CDbl(Calcular("SELECT 'Codigo'=CASE WHEN TMONEDA='01' THEN SUM(NIMPORTE) ELSE SUM(NIMPORTE*NCAMBIO) END From MPAGO Where tCodigoProveedor='" & Prov & "' And tDocumento='" & Docu & "' Group By tMoneda", CnAlmacen))
    VerificaAbonosNC = dAbonos
End Function

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
'    nColumna = ColIndex
'    lAsc = Not lAsc
'    If Sw Then
'       If lAsc Then
'          RsDtemporal.Sort = GrdGrilla.Columns(ColIndex).DataField & " ASC"
'       Else
'          RsDtemporal.Sort = GrdGrilla.Columns(ColIndex).DataField & " DESC"
'       End If
'    Else
'       If lAsc Then
'          RsDocumentoD.Sort = GrdGrilla.Columns(ColIndex).DataField & " ASC"
'       Else
'          RsDocumentoD.Sort = GrdGrilla.Columns(ColIndex).DataField & " DESC"
'       End If
'    End If
End Sub



Private Sub CapturaDatos()
    qCboTipoDocumento = CboTipoDocumento
    qLblPDoc = LblPDoc
    qTxtSerie = TxtSerie
    qTxtCorrelativo = TxtCorrelativo
    qDtpEmision = DtpEmision
    qDtpIngreso = DtpIngreso
    qTxtRUC = TxtRUC
    qTxtNombre = TxtNombre
    qTxtProveedor = TxtProveedor
    'qChkPrograma = ChkPrograma
    'qTxtDiasCredito = TxtDiasCredito
    'qDtpfPrograma = DtpFPrograma
    qCboTipoIngreso = CboTipoIngreso
    qCboMoneda = CboMoneda
    'qCboReferencia = CboReferencia
    'qLblPDocRef = LblPDocRef
    'qTxtSerieRef = TxtSerieRef
    'qTxtCorrelativoRef = TxtCorrelativoRef
    qTxtGlosaDocumento = TxtGlosaDocumento
End Sub



Public Sub ValidaProducto()
    Dim RsValida As Recordset
    If CboArea.BoundText = "000" Then
        Isql = "select distinct a.tCodigoProducto, count(a.tCodigoProducto) as Producto from mKardex a, dDocumento b where fRegistro>'" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "' and a.tCodigoProducto=b.tCodigoProducto and b.tCorrelativo = '" & LblCorrelativo & "' group by a.tCodigoProducto"
        Set RsValida = Lib.OpenRecordset(Isql, CnAlmacen)
        If RsValida.RecordCount > 0 Then
            lValidaProducto = True
        Else
            lValidaProducto = False
        End If
    Else
        Isql = "select distinct a.tCodigoProducto, count(a.tCodigoProducto) as Producto from mKardex a, dDocumento b where fRegistro>'" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "' and a.tCodigoProducto=b.tCodigoProducto and b.tCorrelativo = '" & LblCorrelativo & "' group by a.tCodigoProducto"
        Set RsValida = Lib.OpenRecordset(Isql, CnAlmacen)
        If RsValida.RecordCount > 0 Then
            lValidaProducto = True
        Else
            Isql = "select distinct a.tCodigoProducto, count(a.tCodigoProducto) as Producto from mSubKardex a, dDocumento b where fRegistro>'" & Format(fFechaProcesa, "YYYY/mm/dd HH:mm:ss") & "' and tCodigoSubArea='" & CboArea.BoundText & "' and a.tCodigoProducto=b.tCodigoProducto and b.tCorrelativo = '" & LblCorrelativo & "' group by a.tCodigoProducto"
            Set RsValida = Lib.OpenRecordset(Isql, CnAlmacen)
            If RsValida.RecordCount > 0 Then
                lValidaProducto = True
            Else
                lValidaProducto = False
            End If
        End If
    End If
    Set RsValida = Nothing
End Sub

Private Sub cargatabla()
    Isql = "SELECT " & sDtemporal & ".tCodigoProducto,tProducto.tCodigoBarra As cBarra,tProducto.tDetallado As Producto," & sDtemporal & ".nImpuesto1," & _
           "" & sDtemporal & ".nImpuesto2, " & sDtemporal & ".nImpuesto3, vUnidadMedida.DESCRIPCION AS UnidadMedida, " & sDtemporal & ".nCantidad," & _
           "" & sDtemporal & ".nPrecio, " & sDtemporal & ".nNeto, " & sDtemporal & ".nDescuento, " & sDtemporal & ".nRecargo, " & sDtemporal & ".tUnidadEntrada, " & sDtemporal & ".nItem, " & _
           "" & sDtemporal & ".tCodigoProveedor, " & sDtemporal & ".tDocumento," & sDtemporal & ".nTotal,tProducto.nStockActual," & _
           "" & sDtemporal & ".nPorcentaje1, " & sDtemporal & ".nPorcentaje2, " & sDtemporal & ".nPorcentaje3, " & sDtemporal & ".nPVenta, " & _
           "" & sDtemporal & ".lIncluidoImpuesto1, " & sDtemporal & ".lIncluidoImpuesto2, " & sDtemporal & ".lIncluidoImpuesto3, " & sDtemporal & ".tCorrelativo, " & sDtemporal & ".tEstadoDocumento, " & _
           "" & sDtemporal & ".tCuentacontable, " & sDtemporal & ".tCentroCosto, " & sDtemporal & ".tTipoIngreso , " & sDtemporal & ".tAreaOrigen, " & sDtemporal & ".lAsumirIUE_IT, " & sDtemporal & ".nRetencion1," & _
           "" & sDtemporal & ".nRetencion2," & sDtemporal & ".nRetencion3," & _
           "" & sDtemporal & ".nPRetencion1," & sDtemporal & ".nPRetencion2," & _
           "" & sDtemporal & ".nPRetencion3," & sDtemporal & ".lIncluidoRetencion1," & _
           "" & sDtemporal & ".lIncluidoRetencion2," & sDtemporal & ".lIncluidoRetencion3," & _
           "" & sDtemporal & ".nICE," & sDtemporal & ".nOtrosCargosInafecto" & _
           " FROM " & sDtemporal & " Left Outer Join " & _
           " vUnidadMedida ON " & sDtemporal & ".tUnidadEntrada = vUnidadMedida.Codigo LEFT OUTER JOIN " & _
           " tProducto ON " & sDtemporal & ".tCodigoProducto = tProducto.tCodigoProducto " & _
           " ORDER BY " & sDtemporal & ".tDocumento"
'           Debug.Print Isql
    Set RsDtemporal = Lib.OpenRecordset(Isql, CnAlmacen)
    Set grdGrilla.DataSource = RsDtemporal
End Sub

Function FuncPorcentajeLeyAD(ByVal lTipo As Boolean, ByVal sCorre As String) As Double
    
    Dim LeyAD As String
    
    LeyAD = ""
    If LeyAD = "" Then
        FuncPorcentajeLeyAD = 0
    Else
        If lTipo Then
            FuncPorcentajeLeyAD = Calcular("SELECT sum((dbo." & sDtemporal & ".nPrecio * IsNull(dbo.TPRODUCTO.nPorcentajeLeyAD,0))/100) as Codigo FROM dbo." & sDtemporal & " INNER JOIN dbo.TPRODUCTO ON dbo." & sDtemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
                                      "where dbo.TPRODUCTO.lLeyAD=1", CnAlmacen)
            If FuncPorcentajeLeyAD <> 0 Then
                If MsgBox("¿Documento con " & LeyAD & "?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                    FuncPorcentajeLeyAD = 0
                End If
            End If
        Else
            FuncPorcentajeLeyAD = Calcular("SELECT sum((dbo.DDOCUMENTO.nPrecio * IsNull(dbo.TPRODUCTO.nPorcentajeLeyAD,0))/100) as Codigo FROM dbo.DDOCUMENTO INNER JOIN dbo.TPRODUCTO ON dbo.DDOCUMENTO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
                                      "where dbo.TPRODUCTO.lLeyAD=1 AND dbo.DDOCUMENTO.tCorrelativo='" & sCorre & "'", CnAlmacen)
                                      
            If FuncPorcentajeLeyAD <> 0 Then
                If MsgBox("¿Documento con " & LeyAD & "?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                    FuncPorcentajeLeyAD = 0
                Else
                    Isql = "Exec spSaveLeyADDetails '" & sCorre & "'"
                    Debug.Print Isql
                    CnAlmacen.Execute Isql
                End If
            End If
        End If
    End If
End Function

Function FuncPercepcion(ByVal lTipo As Boolean, ByVal sCorre As String) As Double
        If lTipo Then
            FuncPercepcion = Calcular("SELECT sum((dbo." & sDtemporal & ".nTotal*dbo.TPRODUCTO.nPorcentajeLey)/100) as Codigo FROM dbo." & sDtemporal & " INNER JOIN dbo.TPRODUCTO ON dbo." & sDtemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
                                      "where dbo.TPRODUCTO.lPercepcion=1", CnAlmacen)
            If FuncPercepcion <> 0 Then
                If MsgBox("¿Documento con " & lblLeyD.Caption & "?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                    FuncPercepcion = 0
                End If
            End If
        Else
            FuncPercepcion = Calcular("SELECT sum((dbo.DDOCUMENTO.nTotal*dbo.TPRODUCTO.nPorcentajeLey)/100) as Codigo FROM dbo.DDOCUMENTO INNER JOIN dbo.TPRODUCTO ON dbo.DDOCUMENTO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
                                      "where dbo.TPRODUCTO.lPercepcion=1 AND dbo.DDOCUMENTO.tCorrelativo='" & sCorre & "'", CnAlmacen)
                                      
            If FuncPercepcion <> 0 Then
                If MsgBox("¿Documento con " & lblLeyD.Caption & "?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                    FuncPercepcion = 0
                Else
                    'Guarda la percepcion en el detalle del documento - ddocumento
                    'cambio jesus 2014-09-11
                    Isql = "Exec spSavePerceptionDetails '" & sCorre & "'"
                    Debug.Print Isql
                    CnAlmacen.Execute Isql
                    'fin cambio
                End If
            End If
        End If
End Function

Function RevisaDocExportado(ByVal frdeDocumento As String, ByVal frdeProveedor As String) As Boolean
    If Calcular("select count(tDocumento) as Codigo from mDocumento where tDocumento='" & frdeDocumento & "' and tCodigoProveedor='" & frdeProveedor & "' and tEstadoDocumento<>'04'", CnAlmacen) > 0 Then
        MsgBox "Documeno ya ingresado, verifique ...", vbInformation, sMensaje
        TxtCorrelativo.Text = ""
        TxtSerie.Text = ""
        RevisaDocExportado = True
    Else
        RevisaDocExportado = False
    End If
End Function

Private Sub txtTransporte_KeyPress(KeyAscii As Integer)
    'swPunto = ChrBuscaPunto(txtTransporte.Text)
    If KeyAscii = 13 Then
        If Trim(txtTransporte) = "" Then txtTransporte = "0.00"
        If CboArea.Enabled = True Then
            CboArea.SetFocus
        End If
    Else
        PictureNumero KeyAscii
    End If
End Sub

Private Sub txtTransporte_LostFocus()
    If IsNumeric(txtTransporte.Text) Then
        txtTransporte = Format(txtTransporte, "###,###,##0.00")
    End If
End Sub


