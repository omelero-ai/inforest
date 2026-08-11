VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form frmCtaCteDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Cuenta Corriente"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmCtaCteDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11910
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3240
      TabIndex        =   131
      Top             =   2760
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmCtaCteDetalle.frx":0442
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo XXXX almacenado."
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
         Height          =   210
         Left            =   1185
         TabIndex        =   134
         Top             =   1155
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a XXXX."
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
         Height          =   210
         Left            =   1170
         TabIndex        =   133
         Top             =   870
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmCtaCteDetalle.frx":0655
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Facturación Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   132
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmCtaCteDetalle.frx":0868
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmCtaCteDetalle.frx":0BAA
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a XXXXX......."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   210
         TabIndex        =   135
         Top             =   435
         Width           =   5910
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   75
      TabIndex        =   3
      Top             =   450
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   15478
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pedidos a Facturar"
      TabPicture(0)   =   "frmCtaCteDetalle.frx":0EEC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOrdenCompra"
      Tab(0).Control(1)=   "PrinterFiscalArgentina"
      Tab(0).Control(2)=   "cmdOpcion(15)"
      Tab(0).Control(3)=   "Picture3"
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(7)=   "cmdMovimiento(3)"
      Tab(0).Control(8)=   "cmdMovimiento(2)"
      Tab(0).Control(9)=   "cmdMovimiento(0)"
      Tab(0).Control(10)=   "cmdMovimiento(1)"
      Tab(0).Control(11)=   "cmdOpcion(3)"
      Tab(0).Control(12)=   "cmdOpcion(4)"
      Tab(0).Control(13)=   "cmdOpcion(5)"
      Tab(0).Control(14)=   "Picture1"
      Tab(0).Control(15)=   "Frame1"
      Tab(0).Control(16)=   "grdDetalle1"
      Tab(0).Control(17)=   "grdDetalle2"
      Tab(0).Control(18)=   "Label5(0)"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Generar Guía"
      TabPicture(1)   =   "frmCtaCteDetalle.frx":0F08
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOpcion(11)"
      Tab(1).Control(1)=   "txtMotivoTraslado"
      Tab(1).Control(2)=   "cmdOpcion(16)"
      Tab(1).Control(3)=   "Frame13"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame8"
      Tab(1).Control(6)=   "Picture4"
      Tab(1).Control(7)=   "cmdOpcion(7)"
      Tab(1).Control(8)=   "cmdOpcion(6)"
      Tab(1).Control(9)=   "cmdMovimiento(5)"
      Tab(1).Control(10)=   "cmdMovimiento(4)"
      Tab(1).Control(11)=   "cmdMovimiento(6)"
      Tab(1).Control(12)=   "cmdMovimiento(7)"
      Tab(1).Control(13)=   "Frame7"
      Tab(1).Control(14)=   "Frame4"
      Tab(1).Control(15)=   "Picture2"
      Tab(1).Control(16)=   "grdDetalle3"
      Tab(1).Control(17)=   "grdDetalle4"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Guías a Facturar"
      TabPicture(2)   =   "frmCtaCteDetalle.frx":0F24
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grdDetalle6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "grdDetalle5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame9"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Picture5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdOpcion(10)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdOpcion(9)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdMovimiento(9)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdMovimiento(8)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdMovimiento(10)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdMovimiento(11)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame10"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Frame11"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Frame12"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Picture6"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdOpcion(17)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdOpcion(8)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtOrdenCompra2"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      Begin VB.TextBox txtOrdenCompra2 
         Height          =   405
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   139
         Top             =   6480
         Width           =   1390
      End
      Begin VB.TextBox txtOrdenCompra 
         Height          =   405
         Left            =   -69830
         MaxLength       =   15
         TabIndex        =   137
         Top             =   6480
         Width           =   1390
      End
      Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
         Left            =   -69720
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Motivo Translado"
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
         Index           =   11
         Left            =   -69750
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   7080
         Width           =   1260
      End
      Begin VB.TextBox txtMotivoTraslado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74880
         MaxLength       =   300
         TabIndex        =   129
         Top             =   7320
         Width           =   5050
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         Caption         =   "Anular Guia"
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
         Index           =   8
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":0F40
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   3825
         Width           =   1260
      End
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
         Height          =   615
         Index           =   17
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   7665
         Width           =   1275
      End
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
         Height          =   615
         Index           =   16
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":1134
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   8025
         Width           =   1275
      End
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
         Height          =   615
         Index           =   15
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":1226
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   7665
         Width           =   1275
      End
      Begin VB.Frame Frame13 
         Caption         =   " Tienda "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -74880
         TabIndex        =   118
         Top             =   6225
         Width           =   5055
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Tienda"
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
            Index           =   12
            Left            =   3675
            TabIndex        =   119
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label txtTienda 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   75
            TabIndex        =   120
            Top             =   330
            Width           =   3495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Tipos de Guias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -74880
         TabIndex        =   109
         Top             =   7740
         Width           =   5055
         Begin VB.CommandButton cmdTipoGuia 
            Height          =   555
            Index           =   4
            Left            =   3765
            TabIndex        =   113
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoGuia 
            Height          =   555
            Index           =   3
            Left            =   2550
            TabIndex        =   112
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoGuia 
            Height          =   555
            Index           =   2
            Left            =   1335
            TabIndex        =   111
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoGuia 
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   110
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   97
         Top             =   5940
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   26
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":1318
            Style           =   1  'Graphical
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   24
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":185A
            Style           =   1  'Graphical
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   29
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":1D9C
            Style           =   1  'Graphical
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   28
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":22DE
            Style           =   1  'Graphical
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   27
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":2820
            Style           =   1  'Graphical
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   25
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":2D62
            Style           =   1  'Graphical
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   104
            Top             =   150
            Width           =   1635
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " Cliente a Facturar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   6600
         TabIndex        =   95
         Top             =   7380
         Width           =   5055
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Cliente a Facturar"
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
            Index           =   2
            Left            =   3720
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label txtClienteGuiaFactura 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " Tipos de Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   120
         TabIndex        =   90
         Top             =   7380
         Width           =   5055
         Begin VB.CommandButton cmdTipoDocumento2 
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   94
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento2 
            Height          =   555
            Index           =   2
            Left            =   1335
            TabIndex        =   93
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento2 
            Height          =   555
            Index           =   3
            Left            =   2520
            TabIndex        =   92
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento2 
            Height          =   555
            Index           =   4
            Left            =   3765
            TabIndex        =   91
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Height          =   570
         Left            =   120
         TabIndex        =   87
         Top             =   6540
         Width           =   5055
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total a Pagar :"
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
            Index           =   6
            Left            =   2460
            TabIndex        =   89
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   88
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   11
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":32A4
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2520
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   10
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":382E
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   1815
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   8
         Left            =   5257
         Picture         =   "frmCtaCteDetalle.frx":3DB8
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   420
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   9
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":4342
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1125
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Visualizar Pedido"
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
         Index           =   9
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   4537
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Imp. Guia"
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
         Index           =   10
         Left            =   5250
         Picture         =   "frmCtaCteDetalle.frx":48CC
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   5250
         Width           =   1260
      End
      Begin VB.PictureBox Picture5 
         Height          =   615
         Left            =   6600
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   73
         Top             =   5940
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   31
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":4DFE
            Style           =   1  'Graphical
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   33
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":5340
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   34
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":5882
            Style           =   1  'Graphical
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   35
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":5DC4
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   30
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":6306
            Style           =   1  'Graphical
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   32
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":6848
            Style           =   1  'Graphical
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   80
            Top             =   150
            Width           =   1635
         End
      End
      Begin VB.Frame Frame9 
         Height          =   570
         Left            =   6600
         TabIndex        =   70
         Top             =   6540
         Width           =   5055
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   72
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total a Pagar :"
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
            Index           =   5
            Left            =   2460
            TabIndex        =   71
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   62
         Top             =   5940
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":6D8A
            Style           =   1  'Graphical
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":72CC
            Style           =   1  'Graphical
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":780E
            Style           =   1  'Graphical
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":7D50
            Style           =   1  'Graphical
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":8292
            Style           =   1  'Graphical
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":87D4
            Style           =   1  'Graphical
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   69
            Top             =   150
            Width           =   1635
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Cliente a Facturar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -68400
         TabIndex        =   59
         Top             =   7380
         Width           =   5055
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Cliente a Facturar"
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
            Index           =   0
            Left            =   3720
            TabIndex        =   60
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label txtCliente 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Tipos de Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -74880
         TabIndex        =   54
         Top             =   7380
         Width           =   5055
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   555
            Index           =   2
            Left            =   1335
            TabIndex        =   57
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   555
            Index           =   3
            Left            =   2550
            TabIndex        =   56
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   555
            Index           =   4
            Left            =   3765
            TabIndex        =   55
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   570
         Left            =   -74880
         TabIndex        =   51
         Top             =   6540
         Width           =   5055
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total por Facturar :"
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
            Index           =   3
            Left            =   2460
            TabIndex        =   53
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label txtTotal1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   52
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   3
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":8D16
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2520
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   2
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":92A0
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1815
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   0
         Left            =   -69743
         Picture         =   "frmCtaCteDetalle.frx":982A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   420
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   1
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":9DB4
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1125
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Visualizar Pedido"
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
         Index           =   3
         Left            =   -69750
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3885
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Imp. Cuenta"
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
         Index           =   4
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":A33E
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4575
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         Caption         =   "Anular Pedido"
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
         Index           =   5
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":A870
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5280
         Width           =   1260
      End
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   -68400
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   36
         Top             =   5940
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   7
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":A972
            Style           =   1  'Graphical
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   9
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":AEB4
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   10
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":B3F6
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   11
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":B938
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   6
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":BE7A
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   8
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":C3BC
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   43
            Top             =   150
            Width           =   1635
         End
      End
      Begin VB.Frame Frame1 
         Height          =   570
         Left            =   -68400
         TabIndex        =   33
         Top             =   6540
         Width           =   5055
         Begin VB.Label txtTotal2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   35
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total a Facturar :"
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
            Index           =   2
            Left            =   2580
            TabIndex        =   34
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame Frame8 
         Height          =   570
         Left            =   -68400
         TabIndex        =   30
         Top             =   5640
         Width           =   5055
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total a Generar Guia :"
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
            Index           =   4
            Left            =   2150
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label txtTotal4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   31
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   615
         Left            =   -68400
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   22
         Top             =   5040
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   20
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":C8FE
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   18
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":CE40
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   23
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":D382
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   22
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":D8C4
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   21
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":DE06
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   19
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":E348
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   29
            Top             =   150
            Width           =   1635
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Imp. Cuenta"
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
         Index           =   7
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":E88A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4395
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Visualizar Pedido"
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
         Index           =   6
         Left            =   -69750
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3660
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   5
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":EDBC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1125
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   4
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":F346
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   420
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   6
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":F8D0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1815
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   7
         Left            =   -69750
         Picture         =   "frmCtaCteDetalle.frx":FE5A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Frame Frame7 
         Height          =   570
         Left            =   -74880
         TabIndex        =   13
         Top             =   5640
         Width           =   5055
         Begin VB.Label txtTotal3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3720
            TabIndex        =   15
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total de Pendientes por Generar Guías :"
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
            Index           =   1
            Left            =   1100
            TabIndex        =   14
            Top             =   240
            Width           =   2475
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Datos del Destino y Transportista "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -68400
         TabIndex        =   12
         Top             =   6240
         Width           =   5055
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Vehiculo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   14
            Left            =   3720
            TabIndex        =   123
            Top             =   1275
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   13
            Left            =   3720
            TabIndex        =   121
            Top             =   765
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   3720
            TabIndex        =   116
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label txtVehiculo 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   120
            TabIndex        =   124
            Top             =   1312
            Width           =   3495
         End
         Begin VB.Label txtTransporte 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   120
            TabIndex        =   122
            Top             =   802
            Width           =   3495
         End
         Begin VB.Label txtClienteGuia 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   120
            TabIndex        =   117
            Top             =   292
            Width           =   3495
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   4
         Top             =   5040
         Width           =   5055
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   13
            Left            =   540
            Picture         =   "frmCtaCteDetalle.frx":103E4
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   15
            Left            =   3360
            Picture         =   "frmCtaCteDetalle.frx":10926
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   16
            Left            =   3900
            Picture         =   "frmCtaCteDetalle.frx":10E68
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   17
            Left            =   4440
            Picture         =   "frmCtaCteDetalle.frx":113AA
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   12
            Left            =   0
            Picture         =   "frmCtaCteDetalle.frx":118EC
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   14
            Left            =   1080
            Picture         =   "frmCtaCteDetalle.frx":11E2E
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1680
            TabIndex        =   11
            Top             =   150
            Width           =   1635
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle1 
         Height          =   5520
         Left            =   -74880
         TabIndex        =   105
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9737
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
         Caption         =   "Pedidos por Facturar"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H800000&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle2 
         Height          =   5520
         Left            =   -68400
         TabIndex        =   106
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9737
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
         Caption         =   "Pedidosa a Facturar"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&HFF&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle3 
         Height          =   4620
         Left            =   -74880
         TabIndex        =   107
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   8149
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
         Caption         =   "Pendientes por Generar Guía"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H800000&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle4 
         Height          =   4620
         Left            =   -68400
         TabIndex        =   108
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   8149
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
         Caption         =   "Pedidos a Generar Guía"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&HFF&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle5 
         Height          =   5520
         Left            =   120
         TabIndex        =   114
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9737
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
         Caption         =   "Guias por Facturar"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H800000&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle6 
         Height          =   5520
         Left            =   6600
         TabIndex        =   115
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9737
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
         Caption         =   "Guias a Facturar"
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
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&HFF&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H8000000D&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(57)  =   ":id=37,.fontname=Arial"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Orden Compra"
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
         Index           =   1
         Left            =   5280
         TabIndex        =   138
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Orden Compra"
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
         Index           =   0
         Left            =   -69720
         TabIndex        =   136
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Image imageHash 
      Height          =   615
      Left            =   11880
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label txtCtaCte 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   10035
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Corriente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmCtaCteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Option Explicit
Dim nMontoPedidoFacturar As Double
Dim numeroSerieImpresora As String
Dim codigoImpresora As String
Dim sTemporal1 As String
Dim sSerie As String
Dim sCorrela As String
Dim sPrefijo As String
Dim sImp As String
Dim sTipoEmision As String
Dim sTipoDocumento As String
Dim rstFuente As ADODB.Recordset
Dim RsTipoDocumento As Recordset
Dim RsTipoGuia As Recordset
Dim RsDetalle1 As Recordset
Dim RsDetalle2 As Recordset
Dim RsDetalle3 As Recordset
Dim RsDetalle4 As Recordset
Dim RsDetalle5 As Recordset
Dim RsDetalle6 As Recordset
Dim RsCtaCte As Recordset
Dim RsCliente As Recordset
Dim RsImpresion As Recordset
Dim RsRapido As Recordset
Dim sTemporal As String
Dim sCliente As String
Dim sIdentidad As String
Dim sCtaCte As String
Dim sTienda As String
Dim sTransportista As String
Dim sVehiculo As String
Dim i As Integer
Dim j As Integer

Dim lImprimeAlternativa As Boolean
Dim sUsuarioAutoriza As String
'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tCodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String
Dim lAplicablePedido As Boolean

Dim sComandaInfhotel As String
Dim sPuntoVenta As String
Dim sXML As String

'Por Consumo
Dim sDetalleConsumo As String

'FACTURACION_E_PERU
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim lcodigoHash As Boolean
Dim sImporteLetra As String
'Dim RsImpDocumentoE As New Recordset
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String
Dim TimpresionDolaresDelivery As Boolean
Dim cadenaCodigoHash As String

'----NUEVO
'---- GCAA 12082021
Dim RsNewRptGuia As New dsrNewFormatoGuia
Dim RsPrinter As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          
          Case Is = 0
               Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.cmdOpcion(1).Enabled = True
               frmBusquedaRapida.cmdOpcion(2).Enabled = True
               frmBusquedaRapida.cmdOpcion(3).Enabled = True
               
               Select Case pais
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else  'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
               End Select
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sCliente = sCodigo
                  txtCliente.Caption = sDescrip
                  txtClienteGuia.Caption = sDescrip
                  txtClienteGuiaFactura.Caption = sDescrip
                  Cn.Execute "update TDELIVERY set tCodigoCliente='" & sCodigo & "' where tCodigoDelivery='" & sCtaCte & "'"
               End If
          
          
          Case Is = 1
               If Not IsNumeric(sTemp) Then sTemp = ""
               Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.cmdOpcion(1).Enabled = True
               frmBusquedaRapida.cmdOpcion(2).Enabled = True
               frmBusquedaRapida.cmdOpcion(3).Enabled = True
               
               Select Case pais
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Nit", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
               End Select
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sCliente = sCodigo
                  txtCliente.Caption = sDescrip
                  txtClienteGuia.Caption = sDescrip
                  txtClienteGuiaFactura.Caption = sDescrip
                  Cn.Execute "update TDELIVERY set tCodigoCliente='" & sCodigo & "' where tCodigoDelivery='" & sCtaCte & "'"
               End If
          
          
          Case Is = 2
               Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.cmdOpcion(1).Enabled = True
               frmBusquedaRapida.cmdOpcion(2).Enabled = True
               
               frmBusquedaRapida.cmdOpcion(3).Enabled = True
               
               Select Case pais
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Nit", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
               End Select
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sCliente = sCodigo
                  txtCliente.Caption = sDescrip
                  txtClienteGuia.Caption = sDescrip
                  txtClienteGuiaFactura.Caption = sDescrip
                  Cn.Execute "update TDELIVERY set tCodigoCliente='" & sCodigo & "' where tCodigoDelivery='" & sCtaCte & "'"
               End If
         
         Case Is = 3  ' Visualizar Pedido
            If RsDetalle1.RecordCount = 0 Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            sPedido = RsDetalle1!tCodigoPedido
 
          sTipo = "SinMozo"
   
            frmPedido.cmdOpcion(5).Enabled = False
            frmPedido.cmdOpcion(6).Enabled = False
            frmPedido.cmdOpcion(7).Enabled = False
            frmPedido.Show vbModal
          
         Case Is = 4  ' ReImpresion el Pedido
            If Not (RsDetalle1.EOF Or RsDetalle1.BOF) Then
                Isql = "select * from vCtaCte " & _
                       "WHERE Codigo='" & RsDetalle1!tCodigoPedido & "'"
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                
                If RsImpresion.RecordCount = 0 Then
                   LimpiaRs
                   MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                Else
                    
                    If Supervisor("25") = False Then
                       MsgBox "Clave no permitida", vbExclamation, sMensaje
                       Exit Sub
                    End If
               
                    ImprimeCtaCte RsImpresion
                   
                End If
                LimpiaRs
            End If
                    
         Case Is = 5  ' Revivir el Pedido
                If RsDetalle1.RecordCount <> 0 Then
                   Screen.MousePointer = vbDefault
                   sUsuarioAutoriza = sUsuario
                   If Supervisor("13") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   sUsuarioAutoriza = sVar1
                   Screen.MousePointer = vbHourglass
                   
                   If Periodo(RsDetalle1!fFecha) Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                                
                   Cn.Execute "update MPEDIDO set tEstadoPedido = '01', tClienteCtaCte=''  where tCodigoPedido = '" & RsDetalle1!tCodigoPedido & "'"
                   Cn.Execute "delete from " & sTemporal & " where tCodigoPedido = '" & RsDetalle1!tCodigoPedido & "'"
                   
                   'Actualiza la Cuenta Corriente
                   Cn.Execute "update TDELIVERY set nConsumo = nConsumo - " & RsDetalle1!nVenta & " where tCodigoDelivery ='" & sCtaCte & "'"
                   frmCtaCte.RsCabecera.Requery
                   
                   If lInfhotel Then
                      Dim sComandaInfhotel As String
                      Dim sPuntoVenta As String
                      sComandaInfhotel = Calcular("select tComanda as Codigo From MPEDIDO where tCodigoPedido='" & RsDetalle1!tCodigoPedido & "'", Cn)
                      sPuntoVenta = Calcular("select tPuntoVenta as Codigo From MPEDIDO where tCodigoPedido='" & RsDetalle1!tCodigoPedido & "'", Cn)
                      
                      CnInfhotel.Execute "update MCOMANDA set TESTADO='04' where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                      CnInfhotel.Execute "delete from DCOMANDA where tcomanda='" & sComandaInfhotel & "' and tcodigoitem='100000' and tPuntoVenta='" & sPuntoVenta & "'"
                      CnInfhotel.Execute "delete from WMCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                      CnInfhotel.Execute "delete from WDCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                   End If
                   RsDetalle1.Requery
                   CalculaMonto
                   Screen.MousePointer = vbDefault
                   
                End If
                
         Case Is = 6  ' Visualizar Pedido
            If RsDetalle3.RecordCount = 0 Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            sPedido = RsDetalle3!tCodigoPedido
 
            sTipo = "SinMozo"
            frmPedido.cmdOpcion(5).Enabled = False
            frmPedido.cmdOpcion(6).Enabled = False
            frmPedido.cmdOpcion(7).Enabled = False
            frmPedido.Show vbModal
          
         Case Is = 7  ' ReImpresion el Pedido
            If Not (RsDetalle3.EOF Or RsDetalle3.BOF) Then
                Isql = "select * from vCtaCte " & _
                       "WHERE Codigo='" & RsDetalle3!tCodigoPedido & "'"
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                
                If RsImpresion.RecordCount = 0 Then
                   LimpiaRs
                   MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                Else
                    If Supervisor("25") = False Then
                       MsgBox "Clave no permitida", vbExclamation, sMensaje
                       Exit Sub
                    End If
                    
                    ImprimeCtaCte RsImpresion
                End If
                LimpiaRs
            End If
          
         Case Is = 8  ' Anular Guia de Transporte
                If RsDetalle5.RecordCount <> 0 Then
                   Dim RsElimina As Recordset
                   Dim sMensa As String
                   
                   Isql = "select distinct tGuiaTransporte from DPEDIDO where tCodigoPedido in(select distinct tCodigoPedido from DPEDIDO where tGuiaTransporte='" & RsDetalle5!tGuiaTransporte & "')"
                   Set RsElimina = Lib.OpenRecordset(Isql, Cn)
                   If RsElimina.EOF Then
                      Exit Sub
                   End If
                  
                   sMensa = ""
                   For i = 1 To RsElimina.RecordCount
                       If RsElimina!tGuiaTransporte <> RsDetalle5!tGuiaTransporte Then
                          sMensa = sMensa & RsElimina!tGuiaTransporte & ", "
                       End If
                       RsElimina.MoveNext
                   Next i
                   
                   If MsgBox("Deseas eliminar la Guia " & RsDetalle5!tGuiaTransporte & " y las siguiente(s) Guía(s) Relacionada(s)" & Chr(13) & sMensa, vbExclamation + vbYesNo, sMensaje) <> vbYes Then
                      Exit Sub
                   End If
                   
                   sUsuarioAutoriza = sUsuario
                   If Supervisor("13") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   sUsuarioAutoriza = sVar1
                   Screen.MousePointer = vbHourglass
                   
                   If Periodo(RsDetalle5!fFecha) Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                                                                      
                   Cn.Execute "update MGUIATRANSPORTE set tEstadoGuia = '03', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegistroAnulado=getdate() where tGuiaTransporte in  (select distinct tGuiaTransporte from DPEDIDO where tCodigoPedido in(select distinct tCodigoPedido from DPEDIDO where tGuiaTransporte='" & RsDetalle5!tGuiaTransporte & "'))"
                   Cn.Execute "update DPEDIDO set tGuiaTransporte='' where tGuiaTransporte in (select distinct tGuiaTransporte from DPEDIDO where tCodigoPedido in (select distinct tCodigoPedido from DPEDIDO where tGuiaTransporte='" & RsDetalle5!tGuiaTransporte & "'))"
                   Cn.Execute "update mpedido set testadopedido='04' where tcodigopedido in (select distinct tCodigoPedido from DPEDIDO where tGuiaTransporte='" & RsDetalle5!tGuiaTransporte & "')"
                   'Revierte lo Generado
                   Cn.Execute "delete from " & sTemporal
                   Isql = "Insert into " & sTemporal & _
                          " SELECT ' ' as Estado, MPEDIDO.tCodigoPedido, MPEDIDO.fFecha, Sum(DPEDIDO.nPrecioNeto) AS nNeto, Sum(DPEDIDO.nImpuesto1) AS nImpuesto1, Sum(DPEDIDO.nImpuesto2) AS nImpuesto2, Sum(DPEDIDO.nImpuesto3) AS nImpuesto3, Sum(DPEDIDO.nVenta) AS nVenta, Count(DPEDIDO.tItem) AS nItem, tClienteDelivery, tTienda, tGuiaTransporte " & _
                          "FROM MPEDIDO LEFT JOIN DPEDIDO ON MPEDIDO.tCodigoPedido = DPEDIDO.tCodigoPedido " & _
                          "where MPEDIDO.tClienteCtaCte = '" & sCtaCte & "' and DPEDIDO.tEstadoItem = 'N' and MPEDIDO.tEstadoPedido = '04' and isnull(DPEDIDO.tGuiaTransporte,'')='' " & _
                          "GROUP BY MPEDIDO.fFecha, MPEDIDO.tClienteCtaCte, MPEDIDO.tCodigoPedido, MPEDIDO.tClienteDelivery, tTienda, tGuiaTransporte " & _
                          "Union " & _
                          "SELECT '' as Estado, '' as tCodigoPedido, dbo.MGUIATRANSPORTE.fFecha, SUM(dbo.DGUIATRANSPORTE.nPrecioVenta) AS nPrecioVenta, 0 AS nImpuesto1, 0 AS nImpuesto2, 0 AS nImpuesto3, SUM(dbo.DGUIATRANSPORTE.nVenta) AS nVenta, COUNT(dbo.DGUIATRANSPORTE.tItem) AS tItem, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda , dbo.MGUIATRANSPORTE.tGuiaTransporte " & _
                          "FROM dbo.MGUIATRANSPORTE INNER JOIN dbo.DGUIATRANSPORTE ON dbo.MGUIATRANSPORTE.tGuiaTransporte = dbo.DGUIATRANSPORTE.tGuiaTransporte " & _
                          "where MGUIATRANSPORTE.tCodigoDelivery = '" & sCtaCte & "' and MGUIATRANSPORTE.tEstadoGuia = '01' " & _
                          "GROUP BY dbo.MGUIATRANSPORTE.fFecha, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda, dbo.MGUIATRANSPORTE.tGuiaTransporte "
                    
                   Cn.Execute Isql
                   Cn.Execute "update " & sTemporal & " set Estado='F' where isnull(tGuiaTransporte,'') <> '' "
                   RsDetalle1.Requery
                   RsDetalle2.Requery
                   RsDetalle3.Requery
                   RsDetalle4.Requery
                   RsDetalle5.Requery
                   RsDetalle6.Requery
                   CalculaMonto
                   Screen.MousePointer = vbDefault
              End If
          
         Case Is = 9  ' Visualizar Pedido
            If RsDetalle5.RecordCount = 0 Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            sPedido = RsDetalle5!tGuiaTransporte
            frmGuiaTransporteCorrelativoDetalle.Show vbModal
          
         Case Is = 10  ' ReImpresion Guia
              If Not (RsDetalle5.EOF Or RsDetalle5.BOF) Then
                If lAlmacen And sEmpresa = "100" Then
                    'CASO CORP.FERNANDEZ
                    If BDLink = "1" Then
                        Isql = "select vg.*, isnull(at.UnidadEntrada,'Sin UM') as UM from vGuiaTransporte vg left join " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.vPRODUCTO at on vg.tenlace= at.tCodigoProducto where  vg.tGuiaTransporte ='" & RsDetalle5!tGuiaTransporte & "'"
                    Else
                        Isql = "select vg.*, isnull(at.UnidadEntrada,'Sin UM') as UM from vGuiaTransporte vg left join " & sAlmacenMDB & ".dbo.vPRODUCTO at on vg.tenlace= at.tCodigoProducto where  vg.tGuiaTransporte ='" & RsDetalle5!tGuiaTransporte & "'"
                    End If
                Else
                    '24092018 CLIENTE : DUZKA
                     If sEmpresa = "012" Then
                        Isql = "SELECT vGuiaTransporte.*,ISNULL(MPEDIDO.tObservacion,'') as PedidoObservacion,'Sin UM' as UM FROM [vGuiaTransporte] LEFT JOIN DPEDIDO ON "
                        Isql = Isql & " DPEDIDO.tGuiaTransporte=[vGuiaTransporte].tGuiaTransporte "
                        Isql = Isql & " AND [vGuiaTransporte].CODIGO=DPEDIDO.TCODIGOPRODUCTO "
                        Isql = Isql & " LEFT JOIN MPEDIDO ON MPEDIDO.TCODIGOPEDIDO=DPEDIDO.TCODIGOPEDIDO "
                        Isql = Isql & " where vGuiaTransporte.tGuiaTransporte ='" & RsDetalle5!tGuiaTransporte & "'"
                    Else
                        Isql = "SELECT *,'Sin UM' as UM from vGuiaTransporte where tGuiaTransporte ='" & RsDetalle5!tGuiaTransporte & "'"
                    End If
                End If
                 'Isql = "SELECT * from vGuiaTransporte where tGuiaTransporte ='" & RsDetalle5!tGuiaTransporte & "'"
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            
                If RsImpresion.RecordCount = 0 Then
                    LimpiaRs
                    MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                Else
                    'Configura la Impresora
                    sImp = Calcular("select tImpresora as Codigo from vTipoDocumentoImpresora where tCaja='" & sCaja & "' and Transporte=1 and Prefijo= '" & Mid(RsDetalle5!tGuiaTransporte, 1, 1) & "'", Cn)
                    Printer.FontName = sFont
                    Printer.FontBold = False
                    
                     If lImprimirGuiaFormatoCorp Then
                        Genera4 "R"
                        Dim path As String
                        path = ""

                        RsNewRptGuia.xgaaa = "C:\Documents and Settings\Erick\Mis documentos\Mis imágenes\logo fernandez.jpg"
                        RsNewRptGuia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        RsNewRptGuia.PaperOrientation = crPortrait
                        
                        frmEmite.CRViewer.ViewReport
                        frmEmite.Show vbModal
                    Else
                        ImprimeGuiaTransporte RsImpresion, sEmpresa
                    End If
                    
                    LimpiaRs
                    
                    
                End If
              End If
          
          Case Is = 11 'Motivo Traslado Guia
                    frmKeyBoard.txtResultado.Text = txtMotivoTraslado.Text
                    frmKeyBoard.Show vbModal
                    If wEnter Then
                       txtMotivoTraslado.Text = sDescrip
                    End If
                    wEnter = False
          
          Case Is = 12 ' Tienda / Guia
               sTemporal1 = sTemp
               sTemp = ""
               Isql = "SELECT * from vTienda where tCodigoDelivery='" & sCtaCte & "' Order by Descripcion"
               
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Tienda", 2, "Descripcion", 1600, 2, 0, "", _
                                                               "Direccion", 2, "tDireccion", 5500, 0, 0, "")
               If Not RsDetalle4.EOF Then
                  If Not IsNull(RsDetalle4!descripcion) Then
                     sTemp = Trim(RsDetalle4!descripcion)
                  End If
               End If
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sTienda = sCodigo
                  txtTienda.Caption = sDescrip
               End If
               sTemp = sTemporal1
               
          Case Is = 13 'Transporte
               sTemporal1 = sTemp
               sTemp = ""
               Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.cmdOpcion(1).Enabled = True
               frmBusquedaRapida.cmdOpcion(2).Enabled = True
               frmBusquedaRapida.cmdOpcion(3).Enabled = True
               
               Select Case pais
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Nit", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                           "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                           "Cliente", 2, "Descripcion", 5500, 0, 0, "")
               End Select
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sTransportista = sCodigo
                  txtTransporte.Caption = sDescrip
               End If
               sTemp = sTemporal1
          
          Case Is = 14 ' Vehiculo
               sTemporal1 = sTemp
               sTemp = ""
               Isql = "SELECT *, Marca + ' ' + Placa as Descripcion from vVehiculo where lActivo=1 Order by Marca"
               
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Marca", 2, "Descripcion", 3800, 0, 0, "", _
                                                               "Licencia", 2, "Licencia", 3300, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sVehiculo = sCodigo
                  txtVehiculo.Caption = sDescrip
               End If
               sTemp = sTemporal1
            
          Case Is = 15, 16, 17
               Unload Me
          
          
   End Select
End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
    Dim RsSuma As Recordset
    Dim nRespuesta As Integer
    Dim RsItems As Recordset
    Dim sRapido As String
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim sTotal As Double
    Dim xDescuento As Double
    Dim xVenta As Double
    Dim xTVenta As Double
    Dim xImpuesto1 As Double
    Dim xImpuesto2 As Double
    Dim xImpuesto3 As Double
    Dim xNeto As Double
Dim sxpedido As String
    'FACTURACION_E_PERU
    Dim cadenaCodigoHash As String
    
    lImprimeAlternativa = False
    tAutorizacion = ""
    tCodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    'FACTURACION OFISIS
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    
    Dim oComandoCabeceraOfisis1 As clsComando
    Dim oComandoDetalleOfisis1 As clsComando
    Dim oComandoFirmaDocumentoOfisis1 As clsComando

    '------VALIDA CORREO----------
    Dim sTipoDocum As String
    Dim lValidaEmail As Boolean
    Dim sEmail As String

    Dim rdi As Integer
    Dim RsCantDocumentos As Recordset
    Dim ix As Integer
    Dim xUltimoCorrelativo As String
    Dim RsTotalDocumento As Recordset

    'En blanco Grilla por facturar
    If RsDetalle2.RecordCount = 0 Then
       Exit Sub
    End If
        
    Dim RsCorrela As Recordset
    
    Dim fso1 As Object
    Dim sImporteLetra As String
    
    
    lblPaso1.Visible = False
    lblPaso2.Visible = False
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    FrameFeSpring.Visible = False
    
    
    

   'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    End If
       
    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar < CDbl(txtTotal2.Caption) Then
            MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
            wEnter = False
            Exit Sub
        End If
    End If
    
    'Consistencia Cortesia
    Dim sCortesia As String
    Dim sUsuarioAutoriza As String
    
    sCortesia = ""
    sUsuarioAutoriza = ""
    
    If RsTipoDocumento!TTipoEmision = "00" Then
       If Supervisor("04") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       
       sUsuarioAutoriza = sVar1
       sTemp = ""
       
       Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
       
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
       frmBusquedaRapida.nPredeterm = 1
       
       frmBusquedaRapida.Show vbModal
       
       If wEnter = True Then
          sCortesia = sCodigo
       Else
          Exit Sub
       End If
       
    Else
       If Supervisor("13") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       sUsuarioAutoriza = sVar1
    End If
    
    'validacortesia
    If RsTipoDocumento!TTipoEmision = "00" And sCortesia <> "" Then
        Dim nTotalCortesiaActual As Double
        Dim nTopeCortesia As Double
        Dim nTotalDocActual As Double
        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
        If nTopeCortesia > 0 Then
                'nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and month(fregistro)=month(getdate()) ", Cn)
                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and YEAR(FREGISTRO)=YEAR(GETDATE()) AND month(fregistro)=month(getdate()) ", Cn)
                nTotalDocActual = val(txtTotal2.Caption)
                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                            Exit Sub
                    End If
        End If
    End If

    'POR CONSUMO
    Dim wConsumo As Boolean
    
    wConsumo = False
    sDetalleConsumo = ""
    
    If RsTipoDocumento!TTipoEmision <> "00" Then
       If lConsumo4 = True Then
            nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
            If nRespuesta = vbYes Then
               frmKeyBoard.txtResultado = tTextoConsumo
               frmKeyBoard.Show vbModal
               If sDescrip = "" Or Not wEnter Then
                  MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                  Exit Sub
               End If
               sDetalleConsumo = sDescrip
               wConsumo = True
            ElseIf nRespuesta = vbCancel Then
               Exit Sub
            End If
       End If
    End If
        
        
   
        Dim n1 As String
        Dim ixx As Integer
        Dim Xx As Integer
        grdDetalle2.MoveFirst
        For i = 1 To RsDetalle2.RecordCount
         'For ixx = 0 To grdDetalle2.ApproxCount - 1
                    'grdDetalle2.Row = ixx
            n1 = CStr(grdDetalle2.Columns(1))
         If lbloqueadocconsumo = True And wConsumo = True Then
            Xx = Calcular("select count( round(nPrecioVenta,2))as codigo from dpedido where round(nPrecioVenta,2) = 0 and tCodigoPedido ='" & n1 & "'", Cn)
            If Xx > 0 Then
                MsgBox "Información: No se puede emitir el documento por consumo con bonificaciones (item con precio 0)", vbInformation, sMensaje
                Exit Sub
            End If
            
        End If
        
        If lBloqueoAfectoInafecto And wConsumo Then
            Dim inafecto As Integer
            Dim afecto As Integer
            inafecto = Calcular("select count(*) as codigo from dpedido where tcodigopedido='" & n1 & "' and round(nprecioimpuesto1,2)=0", Cn)
            afecto = Calcular("select count(*) as codigo from dpedido where tcodigopedido='" & n1 & "' and round(nprecioimpuesto1,2)>0", Cn)
            
            If inafecto > 0 And afecto > 0 Then
                MsgBox "Bloqueo de emision: Documento por consumo con afectos e inafectos, pedido: " & n1 & " ", vbInformation, sMensaje
                Exit Sub
            End If
            
        End If
        grdDetalle2.MoveNext
        Next
    
        
        
        
        
    'Consistencia Documento con Cliente
    If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= CDbl(txtTotal2.Caption) Or RsTipoDocumento!Monto = 0) And sCliente = "" Then
       MsgBox "Error: El Documento no tiene Cliente", vbCritical, sMensaje
       Exit Sub
       'imprimedni
    Else
        'imprimedni
        Dim RsTc As ADODB.Recordset
        Set RsTc = New ADODB.Recordset
        Set RsTc = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & RsTipoDocumento!TTipoEmision & "','" & sCliente & "'", Cn)
        If Not (RsTc.EOF Or RsTc.BOF) Then
            RsTc.MoveFirst
            If RsTc.Fields(0) <> "ok" Then
                MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                Exit Sub
            End If
        End If
    End If
                                                                                            
    If Not RsTipoDocumento!Cliente Then
       sCliente = ""
    End If
    
    TimpresionDolaresDelivery = False
        '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCtaCte & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    
    '------VALIDA CORREO -----------------
    sTipoDocum = RsTipoDocumento!TTipoEmision
    lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
    
    If lValidaEmail = True And sCliente <> "" Then
       sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
    
       If sEmail = "" Then
          MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
          Exit Sub
       End If
    End If
    
    
    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True And sCliente <> "" Then
        Dim TempUbigeo As String
        Dim TempUrbaniza As String
        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
            MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
            Exit Sub
        End If
    End If
    '-------------------------------------
                                                                                            
    Screen.MousePointer = vbHourglass
    sRapido = dbTemporal(sCaja, 19, "tDocumento", "nVarChar(20)", _
                                    "tItem", "nVarChar(3)", _
                                    "tCodigoPedido", "nVarChar(10)", _
                                    "tCodigoProducto", "nVarChar(7)", _
                                    "nPrecioNeto", "Float", _
                                    "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                    "nPrecioVenta", "Float", _
                                    "nRecargo", "Float", _
                                    "nDescuento", "Float", _
                                    "nCantidad", "Float", _
                                    "nPrecioOficial", "Float", _
                                    "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                    "nVenta", "Float", "nItem", "nVarChar(3)", "ttipopedido", "nvarchar(2)")
                                                                                                                                                                                                                    
    grdDetalle2.MoveFirst
    For i = 1 To RsDetalle2.RecordCount
        Isql = "insert into " & sRapido & _
               "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem,ttipopedido ) " & _
               "select ' ' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tItem, tTipoPedido From DPEDIDO where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "' and (isnull(tFacturado,'0')='0' Or len(ltrim(tFacturado))=0) and tEstadoItem ='N'"
        Cn.Execute Isql
        Cn.Execute "update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "'"
        grdDetalle2.MoveNext
    Next i
    
    'Diferentes tipos de pedido
    Dim tipPedido As String
    Dim P As Integer
    Dim rstTiposPedido As New ADODB.Recordset
        
    'Distintos tipos de pedidos
    Set rstTiposPedido = Lib.OpenRecordset("select tTipoPedido from " & sRapido & " group by ttipopedido", Cn)
 
    If Not (rstTiposPedido.EOF Or rstTiposPedido.BOF) Then
         rstTiposPedido.MoveFirst
         For P = 1 To rstTiposPedido.RecordCount
                RsTipoDocumento.Requery
                RsTipoDocumento.MoveFirst
                RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
                tipPedido = rstTiposPedido.Fields("ttipopedido")
                Set RsCorrela = Lib.OpenRecordset("select * from " & sRapido & " where ttipopedido='" & tipPedido & "'", Cn)
                RsCorrela.MoveFirst
                For i = 1 To RsCorrela.RecordCount
                    RsCorrela!tItem = Mid("000", 1, 3 - Len(Trim(str(i)))) & Trim(str(i))
                    RsCorrela.MoveNext
                Next i
                                            
                 Select Case pais
                     Case "001" 'Bolivia
                        tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                        tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                        If tAutorizacion <> "" And tDosificacion <> "" Then
                           Else
                            MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                            Exit Sub
                        End If
                     Case "002" 'Ecuador
                            tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
                    
                     Case Else 'Peru
                        tAutorizacion = ""
                        tCodigoControl = ""
                        tDosificacion = ""
                        
                 End Select
                                                     
                 'Verifica y genera la cantidad de documentos
                 If RsTipoDocumento!tFormulario <> "01" And Not wConsumo Then
                        Dim xDocumento(50) As String
                        Dim nTotal As Integer
                        nTotal = 1
                        
                        'Factura
                        'Genera y Actualiza los Numero de Documento
                        sSerie = RsTipoDocumento!tSerie
                        sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                        sPrefijo = RsTipoDocumento!prefijo
                        sTipoDocumento = RsTipoDocumento!TTipoEmision
                        sImp = RsTipoDocumento!timpresora
                        sDocumento = sPrefijo & sSerie & sCorrela
                        sResumen = RsTipoDocumento!lResumen
                        xDocumento(nTotal) = sDocumento
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Set RsRapido = Lib.OpenRecordset("select * from " & sRapido & " where ttipopedido='" & tipPedido & "'", Cn)
                        RsRapido.MoveFirst
                        i = 1
                        sTotal = 0
                           
                           
                        Do While Not RsRapido.EOF
                              sTotal = sTotal + 1
                              If i >= nFItem + 1 Then
                                 Isql = "Insert into MDOCUMENTO " & _
                                        "     ( tDocumento, tTipoDocumento, tCortesia, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro, fdiaContable, lImpresionMonedaExtranjera) " & _
                                        "Values(   '" & sDocumento & "', " _
                                                & "'" & sTipoDocumento & "', " _
                                                & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                                & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                                & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                                & "'" & sCaja & "', " _
                                                & "'" & sSalon & "', " _
                                                & "'" & sTurno & "', " _
                                                & "'" & sUsuario & "', " _
                                                & "'" & sUsuarioAutoriza & "', " _
                                                & "'" & IIf(wConsumo, "C", "D") & "', " _
                                                & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                                & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " ) "
                                 Cn.Execute Isql
                                 
                                 'Calcula el total de la cabecera
                                 Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento * nCantidad),0) as nDescuento " & _
                                                                " from " & sRapido & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
                                                                
                                 xDescuento = RsSuma!nDescuento
                                 xVenta = RsSuma!nVenta
                                 xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                                 xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                                 xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                                 xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                                                                
                                Select Case pais
                                    Case "001"
                                        tCodigoControl = devuelveCodigoControl(sCaja, Mid(sDocumento, 7, 9), tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                                End Select
                                                                
                                 'Actualiza el Documento con el Temporal
                                 Isql = "Update MDOCUMENTO set nNeto= " & xNeto & ", " & _
                                                              "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                              "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                              "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                              "nRecargo = 0, " & _
                                                              "nDescuento = " & xDescuento & ", " & _
                                                              "nVenta = " & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " & _
                                                              "nPrecioOficial = " & RsSuma!nVenta & ", " & _
                                                              "tautorizacion = '" & tAutorizacion & "' , " & _
                                                              "tcodigocontrol = '" & tCodigoControl & "' , " & _
                                                              "tordenCompra='" & txtOrdenCompra.Text & "'," & _
                                                              " lreplica=1  where tDocumento = '" & sDocumento & "'"
                                 Cn.Execute Isql
                                 
                                 
                                 'Inserta todos los items del temporal en el DETALLE
                                 Isql = "Insert into DDOCUMENTO " & _
                                        "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                        "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                                        "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                        "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sRapido & " where tDocumento ='" & sDocumento & "'"
                                 Cn.Execute Isql
                                 
                                 
                                 'SACA LA VUELTA 06082021
                                  'aqui 06082021 GCAA
               
                                   Dim oComandox As clsComando
                                   Set oComandox = New clsComando
                    
                                   Set oComandox = New clsComando
                                    If Not oComandox.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                    End If
                                    
                                    oComandox.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento
                                    
                                    If Not oComandox.GetParamOK Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                    End If
                                
                                    If Not oComandox.ExecSP Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                     End If
                                   'fin 06082021 GCAA
               
                                 
                                 
                                 'Actualiza Base de Datos Detalle del Pedido
                                 'Genera y Actualiza los Numero de Documento
                                 RsTipoDocumento.Requery
                                 RsTipoDocumento.MoveFirst
                                 RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
                                 If RsTipoDocumento.EOF Then
                                    MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
                                    Exit Sub
                                 End If
                                         
                                 sSerie = RsTipoDocumento!tSerie
                                 sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                                 sPrefijo = RsTipoDocumento!prefijo
                                 sTipoDocumento = RsTipoDocumento!TTipoEmision
                                                                 
                                 sImp = RsTipoDocumento!timpresora
                                 sDocumento = sPrefijo & sSerie & sCorrela
                                 
                                 sResumen = RsTipoDocumento!lResumen
                                 nTotal = nTotal + 1
                                 
                                 xDocumento(nTotal) = sDocumento
                                 Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                 
                                 i = 2
                                 
                                 RsRapido!tDocumento = sDocumento
                                 RsRapido.Update
                                 
                                 sPedido = IIf(IsNull(RsRapido!tCodigoPedido), " ", RsRapido!tCodigoPedido)
                                 
                              Else
                                 RsRapido!tDocumento = sDocumento
                                 RsRapido.Update
                                 sPedido = IIf(IsNull(RsRapido!tCodigoPedido), " ", RsRapido!tCodigoPedido)
                                 i = i + 1
                              End If
                                 'GCAA ACTUALIZA ORDEN DE COMPRA
                            Isql = "update MPEDIDO set MPEDIDO.tOrdenCompra = '" & txtOrdenCompra.Text & "' where MPEDIDO.tCodigoPedido = '" & sPedido & "'"
                            Cn.Execute Isql
                           
                              RsRapido.MoveNext
                              
                        Loop
                                  
                                  
                            Isql = "Insert into MDOCUMENTO " & _
                                   "     ( tDocumento, tTipoDocumento, tCortesia, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro, fdiacontable, lImpresionMonedaExtranjera) " & _
                                   "Values(   '" & sDocumento & "', " _
                                           & "'" & sTipoDocumento & "', " _
                                           & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                           & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                           & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                           & "'" & sCaja & "', " _
                                           & "'" & sSalon & "', " _
                                           & "'" & sTurno & "', " _
                                           & "'" & sUsuario & "', " _
                                           & "'" & sUsuarioAutoriza & "', " _
                                           & "'" & IIf(wConsumo, "C", "D") & "', " _
                                           & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                           & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " ) "
                                 
                            Cn.Execute Isql
                           
                           'Calcula el total de la cabecera
                           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento * nCantidad),0) as nDescuento " & _
                                                          " from " & sRapido & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
                                                                                                         
                           xVenta = RsSuma!nVenta
                           xDescuento = RsSuma!nDescuento
                           
                           xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                           xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                           
                           xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                           xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                           
                            Select Case pais
                               Case "001"
                                    tCodigoControl = devuelveCodigoControl(sCaja, Mid(sDocumento, 7, 9), tAutorizacion, tDosificacion, sCliente, xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3)
                            End Select
                                                  
                           'Actualiza el Documento con el Temporal
                           Isql = "Update MDOCUMENTO set nNeto= " & xNeto & ", " & _
                                                        "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                        "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                        "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                        "nRecargo = 0, " & _
                                                        "nDescuento = " & xDescuento & ", " & _
                                                        "tautorizacion='" & tAutorizacion & "', " & _
                                                        "tcodigocontrol='" & tCodigoControl & "'," & _
                                                        "nVenta = " & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " & _
                                                        "nPrecioOficial = " & RsSuma!nVenta & " , " & _
                                                        "tordenCompra='" & txtOrdenCompra.Text & "' " & _
                                                        " ,lreplica=1  where tDocumento = '" & sDocumento & "'"
                           Cn.Execute Isql
                                                                  
                           'Inserta todos los items del temporal en el DETALLE
                           Isql = "Insert into DDOCUMENTO " & _
                                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                                  "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sRapido & " where tDocumento ='" & sDocumento & "'"
                           Cn.Execute Isql
                           
                           'SACA LA VUELTA 06082021
                                  'aqui 06082021 GCAA
               
                                   Dim oComandoxx As clsComando
                                   Set oComandoxx = New clsComando
                    
                                   Set oComandoxx = New clsComando
                                    If Not oComandoxx.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                    End If
                                    
                                    oComandoxx.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento
                                    
                                    If Not oComandoxx.GetParamOK Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                    End If
                                
                                    If Not oComandoxx.ExecSP Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                     End If
                            'fin 06082021 GCAA
                           
                            If lPagoAntesImpresion Then
                                For i = 1 To nTotal
                                    wEnter = True
                                    
                                    Set RsTotalDocumento = Lib.OpenRecordset("select nventa from MDOCUMENTO where tDocumento ='" & xDocumento(i) & "'", Cn)
                                    nCargo = Round(RsTotalDocumento!nVenta, 2)
                                    sDocumento = xDocumento(i)
                                    sFormulario = ""
                                    sTipo = ""
                                    frmPago.Show vbModal
                                    
                                    If wEnter = False Then
                                        grdDetalle2.MoveFirst
                                        For ix = 1 To RsDetalle2.RecordCount
                                            Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                            Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                            For rdi = 0 To RsCantDocumentos.RecordCount - 1
                                                Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                RsCantDocumentos.MoveNext
                                            Next rdi
                                            RsDetalle2.MoveNext
                                        Next ix

                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Exit Sub
                                    End If
                                Next i
                            End If
                           
                         
                            'Actualiza Base de Datos Detalle del Pedido
                            Isql = "UPDATE DPEDIDO SET DPEDIDO.tFacturado = 'F', DPEDIDO.tDocumento = " & sRapido & ".tDocumento " & _
                                   "FROM  DPEDIDO INNER JOIN " & sRapido & " ON DPEDIDO.tCodigoPedido = " & sRapido & ".tCodigoPedido AND DPEDIDO.tItem = " & sRapido & ".nItem "
                            Cn.Execute Isql
                            
                            Isql = "UPDATE DPEDIDO SET DPEDIDO.tFacturado = 'C', DPEDIDO.tDocumento = " & sRapido & ".tDocumento " & _
                                   "FROM DPEDIDO INNER JOIN " & sRapido & " ON DPEDIDO.tCodigoPedido = " & sRapido & ".tCodigoPedido AND substring(DPEDIDO.tDocumento,1,1)='0'"
                            Cn.Execute Isql
                             
                            
                            'PARA NO FISCALES
                            Dim RST As New ADODB.Recordset
                            Set RST = Lib.OpenRecordset("select tdocumento from " & sRapido & " order by tdocumento", Cn)
                            If Not (RST.EOF Or RST.BOF) Then
                                 RST.MoveFirst
                                 Do While Not RST.EOF
                                     Isql = "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tDocumento ='" & RST!tDocumento & "'"
                                     Cn.Execute Isql
                                     RST.MoveNext
                                 Loop
                            End If

                           For i = 1 To nTotal
                           
                                 'Imprime Documentos
                                 If wConsumo = False And lDescripcionAlternativa = True Then
                                    If validaImpresionAlternativa(sDocumento) = False Then
                                               If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                                       lImprimeAlternativa = True
                                               End If
                                    End If
                                 End If
                                 '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                                 ' ELDCQ 15/11/2017
                                 If lImprimeAlternativa = False Then
                                        If lDocumentoAgrupado Then
                                            Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',1"
                                            'FACTURACION_E_PERU
                                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',2"
                                        Else
                                            Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',3"
                                            'FACTURACION_E_PERU
                                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',4"
                                        End If
                                Else
                                        If lDocumentoAgrupado Then
                                            Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',5"
                                            'FACTURACION_E_PERU
                                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',6"
                                        Else
                                            Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',7"
                                            'FACTURACION_E_PERU
                                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',8"
                                        End If
                                End If
                                '-------------------------------------------------------------------------------------------------------------
                                
                                'FACTURACION_E_PERU
                                'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                                xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & xDocumento(i) & "')", Cn)
                                xImpresioDE = Mid(sDocumento, 1, 1)
                                '---------------------------------------
                          
                                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                             
                                'impresion imagen
                                Set rstFuente = New ADODB.Recordset
                                imageCab.Picture = Nothing
                                imagepIE.Picture = Nothing
                                Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                                imageCab.DataField = "foto"
                                Set imageCab.DataSource = rstFuente
                                imagepIE.DataField = "fotoPie"
                                Set imagepIE.DataSource = rstFuente
        
    
                             
                               If RsImpresion.RecordCount = 0 Then
                                  LimpiaRs
                                  MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                               Else
                                  'SUNAT
                                  numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                                  codigoImpresora = sImp
                                  'SUNAT
                                  Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & xDocumento(i) & "' "
    
    
                                    'FACTURACION_E_PERU
                                    If pais = "000" Then
                                       If lFacturacionE Then
                                            If lFEOfisis Then 'OFISIS
                                                     '----CABECERA
                                                    Set oComandoCabeceraOfisis = New clsComando
                                                    If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                         Set oComandoCabeceraOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, xDocumento(i)
                                    
                                                    If Not oComandoCabeceraOfisis.GetParamOK Then
                                                         Set oComandoCabeceraOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoCabeceraOfisis.ExecSP Then
                                                         Set oComandoCabeceraOfisis = Nothing
                                                         Exit Sub
                                                    End If

                                                    '----FIRMA DOCUMENTO OFISIS
                                                    If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                        Set oComandoFirmaDocumentoOfisis = New clsComando
                                                        If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                             Set oComandoFirmaDocumentoOfisis = Nothing
                                                             Exit Sub
                                                        End If
                                                        oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, xDocumento(i)
                                    
                                                        If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                             Set oComandoFirmaDocumentoOfisis = Nothing
                                                             Exit Sub
                                                        End If
                                                        If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                             Set oComandoFirmaDocumentoOfisis = Nothing
                                                             Exit Sub
                                                        End If
                                                        
                                                        'VALIDAR RESPUESTA CODIGO DE BARRA
                                                        fDocumento = Mid(xDocumento(i), 1, 1) + Mid(xDocumento(i), 4, 3) + "-" + CStr(CLng(Mid(xDocumento(i), 8, 8)))
                                                        Sleep 3000
                                                        If lImpresionCodigoBarras Then
                                                            imageHash.DataField = "foto"
                                                            Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                            Set imageHash.DataSource = RsCodigoHash
                                                            
                                                        ElseIf lQRFE Then
                                                            Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                        Else
                                                            Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                            If RscadenaCodigoHash.RecordCount > 0 Then
                                                                cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                            End If
                                                            'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                        End If
                                                    End If
                                            
                                            ElseIf lFESpring Then
                                            
                                            ElseIf lFECarbajal Then
                                                Label4.Caption = "   Proceso de envio de documento a InfoFact......."
                                                lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                                lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                                    FrameFeSpring.Visible = True
                                                    lblPaso1.Visible = True
                                                    lblPaso2.Visible = True
                                                    imgProceso(0).Visible = False
                                                    imgProceso(1).Visible = False
                                                    imgProceso(2).Visible = False
                                                    imgProceso(3).Visible = False
                                                    Sleep 1000
                                                    If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                                            xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                            xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                            grdDetalle2.MoveFirst
                                                            For ix = 1 To RsDetalle2.RecordCount
                                                                Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                                                Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                                                Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                RsDetalle2.MoveNext
                                                            Next ix
                                                            imgProceso(2).Visible = True
                                                            imgProceso(3).Visible = True
                                                            Sleep 1000
                                                            FrameFeSpring.Visible = False
                                                            Exit Sub
                                                     End If
                                                     imgProceso(0).Visible = True
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                                             Else
                                                                 cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                                             End If
                                                         End If
                                                     End If
                                                     imgProceso(1).Visible = True
                                                     Sleep 1000
                                                     FrameFeSpring.Visible = False
                                                End If
                                            ElseIf lFEpape Then
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                ' Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                             Else
                                                                 cadenaCodigoHash = PapeMatricial
                                                             End If
                                                         End If
                                                     End If
                                                End If
                                            Else 'INFOFACT
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    If Not INSERTAFE(xDocumento(i), "", 1, "") Then '----CABECERA
                                                        'Exit Sub
                                                    End If
'                                                    If RsImpresion!Ruc <> "" Then
'                                                        If Not INSERTAFE(xDocumento(i), "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                            Exit Sub
'                                                        End If
'                                                    End If
                                                    'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                    fDocumento = Mid(xDocumento(i), 1, 1) + Mid(xDocumento(i), 4, 3) + Mid(xDocumento(i), 8, 8)
                                                    If tCodigoFE = "000" Then
                                                        If lQRFE Then
                                                            Set imageHash.Picture = LoadPicture(ImagenQR(xDocumento(i)))
                                                        Else
                                                            If lImpresionCodigoBarras Then
                                                                Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, xDocumento(i)))
                                                            Else
                                                                cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, xDocumento(i))
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                    
                                       End If
                                                           
                                    End If
                                    '---------------------------------------

                                    'Configura la Impresora
                                    Imprimir (sImp)
                                    Printer.FontName = sFont
                                    Printer.FontBold = False
                                                 
                                    If wConsumo Then
                                       If RsTipoDocumento!tFormulario = "01" Then
                                          If RsTipoDocumento!Cliente Then
                                             ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                             NFactura = sCorrela
                                          Else
                                             ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                          End If
                                       ElseIf RsTipoDocumento!tFormulario = "02" Then
                                          If RsTipoDocumento!Cliente Then
                                             ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                             NFactura = sCorrela
                                          Else
                                             ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                          End If
                                       Else
                                          If RsTipoDocumento!Cliente Then
                                             ImprimeFacturaVariableConsumo RsImpresion, sDetalleConsumo, sEmpresa
                                             NFactura = sCorrela
                                          End If
                                       End If
                                       
                                    Else
                                       'TICKET
                                       If RsTipoDocumento!tFormulario = "01" Then
                                          If RsTipoDocumento!Cliente Then
                                    
                                                 If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                     ImprimeFacturaElectronica RsImpresion, Nothing, "", Nothing, Nothing, "", TimpresionDolaresDelivery
                                                 Else
                                                     ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                                 End If
                                             
                                             NFactura = sCorrela
                                          ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                             If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                                ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                             End If
                                          Else
                                    
                                                 If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                     ImprimeBoletaElectronica RsImpresion, Nothing, "", Nothing, Nothing, "", TimpresionDolaresDelivery
                                                 Else
                                                     ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                                 End If
                                          End If
                                          
                                       'VARIABLE
                                       ElseIf RsTipoDocumento!tFormulario = "02" Then
                                          If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                                                 'FACTURACION_E_PERU
                                                 'FORMATO A4
                                                     If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                                         ImprimeFormatoA xDocumento(i)
                                                         Kill App.path & "\fact.bmp"
                                                     Else
                                                         ImprimeFormatoA xDocumento(i)
                                                     End If
                                          ElseIf lFacturacionE And lFECarbajal Then
                                                    ImprimeFormatoA xDocumento(i)
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                                        Kill App.path & "\fact.bmp"
                                                    End If
                                          Else
                                                     If RsTipoDocumento!TTipoEmision = "01" Then
                                                        ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                                        NFactura = sCorrela
                                                     ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                                        If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                                           ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                                        End If
                                                     Else
                                                        ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocum
                                                     End If
                                          End If
                                          
                                       'TICKET VARIABLE
                                       Else
                                          If RsTipoDocumento!Cliente Then
                                             ImprimeFacturaVariable RsImpresion, sEmpresa
                                             NFactura = sCorrela
                                          ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                             If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                                ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                             End If
                                          Else
                                             'ImprimeBoletaN RsImpresion, sEmpresa
                                          End If
                                       End If
                                       
                                    End If
                                    LimpiaRs
                               End If
                                
'                                If pais = "002" Then
'                                   sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                                   GeneraFacturaElectronica sXML, sDocumento
'                                End If
                                If pais = "002" And lFEEcuador = False Then
                                   sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                                   GeneraFacturaElectronica sXML, sDocumento
                                End If
                                
                                If lFEEcuador Then
                                 If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                                     MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                                 End If
                                End If
                                                   
                           Next i

                    '------------------- FORMATO TICKET
                    Else
                           
                           sSerie = RsTipoDocumento!tSerie
                           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                           
                           sPrefijo = RsTipoDocumento!prefijo
                           sTipoDocumento = RsTipoDocumento!TTipoEmision
                           sImp = RsTipoDocumento!timpresora
                           sDocumento = sPrefijo & sSerie & sCorrela
                           sResumen = RsTipoDocumento!lResumen
                           
                           'Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                  
                           Dim xTempTipPedido As String
                                  
                           'Inserta Detalle de Documento en un temporal
                           grdDetalle2.MoveFirst
                           For i = 1 To RsDetalle2.RecordCount
                               sxpedido = grdDetalle2.Columns(1).Text
                                    'GCAA ACTUALIZA ORDEN DE COMPRA
                                Cn.Execute "update MPEDIDO set tOrdenCompra = '" & txtOrdenCompra.Text & "' where tCodigoPedido = '" & sxpedido & "'"
                           
                           
                               xTempTipPedido = Calcular("Select tTipoPedido as Codigo From MPEDIDO where tCodigoPedido='" & grdDetalle2.Columns(1).Text & "'", Cn)
                              
                               If xTempTipPedido = tipPedido Then
                                  Cn.Execute "update DPEDIDO set tFacturado = 'F', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "' and (Len(lTrim(tFacturado)) = 0 Or IsNull(tFacturado,'0')='0') and tEstadoItem ='N'"
                                  Cn.Execute "update DPEDIDO set tFacturado = 'C', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "' and tEstadoItem ='N' and substring(tDocumento,1,1)='0'"
                                  Cn.Execute "update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "'"
                               End If
                               grdDetalle2.MoveNext
                           Next i
     
                           'PARA NO FISCALES
                           Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and tDocumento ='" & sDocumento & "'"
                                  
                           
                            Isql = "Insert into DDOCUMENTO " & _
                                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nley1 ) " & _
                                  "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  " " & sRapido & ".nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, (case when vproducto.lley1=1 then ncantidad*(" & sValorLey1 & ") else 0 end ) From " & sRapido & " inner join vproducto on  " & sRapido & ".tcodigoproducto= vproducto.codigo where ttipopedido='" & tipPedido & "'"
                           Cn.Execute Isql
                           
                           
                            'SACA LA VUELTA 06082021
                            
                            'SACA LA VUELTA 06082021
                                  'aqui 06082021 GCAA
               
                                   Dim oComandoxxx As clsComando
                                   Set oComandoxxx = New clsComando
                    
                                   Set oComandoxxx = New clsComando
                                    If Not oComandoxxx.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                                       Set oComandoxxx = Nothing
                                       Exit Sub
                                    End If
                                    
                                    oComandoxxx.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento
                                    
                                    If Not oComandoxxx.GetParamOK Then
                                       Set oComandoxxx = Nothing
                                       Exit Sub
                                    End If
                                
                                    If Not oComandoxxx.ExecSP Then
                                       Set oComandoxxx = Nothing
                                       Exit Sub
                                     End If
                            'fin 06082021 GCAA
                            
                           'Calcula el total de la cabecera
                           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento * nCantidad),0) as nDescuento, sum(case when vproducto.lley1=1 then ncantidad*(" & sValorLey1 & ") else 0 end )  as ValorLey1 " & _
                                                          " from " & sRapido & " inner join vproducto on  " & sRapido & ".tcodigoproducto= vproducto.codigo where ttipopedido='" & tipPedido & "'  group by tDocumento", Cn)
                                                                                                                
                           xDescuento = RsSuma!nDescuento
                           xVenta = RsSuma!nVenta + RsSuma!ValorLey1
                           xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                           xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                           xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                           xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                                                         
                           Select Case pais
                                Case "001"
                                    tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                           End Select
                           
                           'Inserta el Documento
                           Isql = "Insert into MDOCUMENTO " & _
                                  "     ( tDocumento, tTipoDocumento, tCortesia, tcodigoCliente, tEstadoDocumento, tCaja, tTurno, nNeto, nDescuento, nRecargo, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, tSalon, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro,tautorizacion,tcodigocontrol, fDiaContable, lImpresionMonedaExtranjera, nley1,tordenCompra) " & _
                                  "Values(   '" & sDocumento & "', " _
                                          & "'" & sTipoDocumento & "', " _
                                          & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                          & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                          & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                          & "'" & sCaja & "', " _
                                          & "'" & sTurno & "', " _
                                          & xNeto & ", " _
                                          & xDescuento & ", " _
                                          & "0 , " _
                                          & RsSuma!nVenta & ", " _
                                          & xImpuesto1 & ", " _
                                          & xImpuesto2 & ", " _
                                          & xImpuesto3 & ", " _
                                          & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " _
                                          & "'" & sSalon & "', " _
                                          & "'" & sUsuario & "', " _
                                          & "'" & sUsuarioAutoriza & "', " _
                                          & "'" & IIf(wConsumo, "C", "D") & "', " _
                                          & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                          & " getdate(),'" & tAutorizacion & "','" & tCodigoControl & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & ", " & RsSuma!ValorLey1 & ", '" & txtOrdenCompra.Text & "' ) "
                                          
                           Cn.Execute Isql

                   
                           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                           
                           If lPagoAntesImpresion And sTipoDocumento <> "00" Then
                                  Set RsTotalDocumento = Lib.OpenRecordset("select nventa from MDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                                  nCargo = Round(RsTotalDocumento!nVenta, 2)
                                  sFormulario = ""
                                  sTipo = ""
                                  frmPago.Show vbModal
                                  
                                  If wEnter = False Then
                                        grdDetalle2.MoveFirst
                                        For ix = 1 To RsDetalle2.RecordCount
                                            Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                            Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"

                                            Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"

                                            RsDetalle2.MoveNext
                                        Next ix

                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Exit Sub
                                  End If
                           End If
                           
                    
                           
                             '-----------------------
                             If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not FacturarTCPIP(2, sDocumento, 0) Then
                                    grdDetalle2.MoveFirst
                                    For ix = 1 To RsDetalle2.RecordCount
                                        Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                        Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        RsDetalle2.MoveNext
                                    Next ix
                                    
                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        
                                    Exit Sub
                                End If
                             End If
                             '------------------------
                                 
                            'Imprime Documentos
                            If wConsumo = False And lDescripcionAlternativa = True Then
                                        If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                                lImprimeAlternativa = True
                                        End If
                            End If
                             '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                             ' ELDCQ 15/11/2017
                             If lImprimeAlternativa = False Then
                                If lDocumentoAgrupado Then
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
                                Else
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                                End If
                            Else
                                If lDocumentoAgrupado Then
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
                                Else
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                                End If
                            End If
                            '---------------------------------------------------------------------------------------------------------------
                            'FACTURACION_E_PERU
                            'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                            xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                            xImpresioDE = Mid(sDocumento, 1, 1)
                            '---------------------------------------
                            
                            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                            
                            'impresion imagen
                            Set rstFuente = New ADODB.Recordset
                            imageCab.Picture = Nothing
                            imagepIE.Picture = Nothing
                            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                            imageCab.DataField = "foto"
                            Set imageCab.DataSource = rstFuente
                            imagepIE.DataField = "fotoPie"
                            Set imagepIE.DataSource = rstFuente
                                
                            
                            If RsImpresion.RecordCount = 0 Then
                                 LimpiaRs
                                 MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                            Else
                                 'SUNAT
                                 numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                                 codigoImpresora = sImp
                                 'SUNAT
                                 Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "
                                  
                                  
                                   'FACTURACION_E_PERU
                                   If pais = "000" Then
                                       If lFacturacionE Then
                                            If lFEOfisis Then 'OFISIS
                                                    '----CABECERA
                                                    Set oComandoCabeceraOfisis1 = New clsComando
                                                    If Not oComandoCabeceraOfisis1.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                         Set oComandoCabeceraOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoCabeceraOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                                   
                                                    If Not oComandoCabeceraOfisis1.GetParamOK Then
                                                         Set oComandoCabeceraOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoCabeceraOfisis1.ExecSP Then
                                                         Set oComandoCabeceraOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    
                                                    '----FIRMA DOCUMENTO OFISIS
                                                    If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                        Set oComandoFirmaDocumentoOfisis1 = New clsComando
                                                        If Not oComandoFirmaDocumentoOfisis1.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                             Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                             Exit Sub
                                                        End If
                                                        oComandoFirmaDocumentoOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                                   
                                                        If Not oComandoFirmaDocumentoOfisis1.GetParamOK Then
                                                             Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                             Exit Sub
                                                        End If
                                                        If Not oComandoFirmaDocumentoOfisis1.ExecSP Then
                                                             Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                             Exit Sub
                                                        End If
                                                        
                                                        'VALIDAR RESPUESTA CODIGO DE BARRA
                                                        fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                        Sleep 3000
                                                        If lImpresionCodigoBarras Then
                                                            imageHash.DataField = "foto"
                                                            Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                            Set imageHash.DataSource = RsCodigoHash
                                                            
                                                        ElseIf lQRFE Then
                                                            Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                        Else
                                                            Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                            If RscadenaCodigoHash.RecordCount > 0 Then
                                                                cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                            End If

                                                            'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                        End If
                                                    End If
                                                    
                                            ElseIf lFESpring Then
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                                                                          
                                                     frmMensajeFeSpring.sDocumento = sDocumento
                                                     frmMensajeFeSpring.oVenta = 3 ' 3: "Formulario Cta.Corriente"
                                                     frmMensajeFeSpring.Show vbModal
                                                     If frmMensajeFeSpring.lEnvio = False Then
                                                         Exit Sub
                                                     End If
                                                    
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                            If frmMensajeFeSpring.lQrInf Then
                                                                Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                            Else
                                                                Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                                            End If
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 
                                                             Else
                                                                 cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                                             End If
                                                         End If
                                                     End If
                                                 End If
                                              
                                            ElseIf lFECarbajal Then
                                                Label4.Caption = "   Proceso de envio de documento a InfoFact......."
                                                lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                                lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                                    FrameFeSpring.Visible = True
                                                    lblPaso1.Visible = True
                                                    lblPaso2.Visible = True
                                                    imgProceso(0).Visible = False
                                                    imgProceso(1).Visible = False
                                                    imgProceso(2).Visible = False
                                                    imgProceso(3).Visible = False
                                                    Sleep 1000
                                                    If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                                            grdDetalle2.MoveFirst
                                                            For ix = 1 To RsDetalle2.RecordCount
                                                                Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                                                Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                                                Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                RsDetalle2.MoveNext
                                                            Next ix
                                                            xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                            xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                            imgProceso(2).Visible = True
                                                            imgProceso(3).Visible = True
                                                            Sleep 1000
                                                            FrameFeSpring.Visible = False
                                                            Exit Sub
                                                     End If
                                                     imgProceso(0).Visible = True
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                                             Else
                                                                 cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                                             End If
                                                         End If
                                                     End If
                                                     imgProceso(1).Visible = True
                                                     Sleep 1000
                                                     FrameFeSpring.Visible = False
                                                End If
                                              
                                            ElseIf lFEpape Then
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                             Else
                                                                 cadenaCodigoHash = PapeMatricial
                                                             End If
                                                         End If
                                                     End If
                                                End If
                                            ElseIf lFEBiz Then
                                                    If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                        If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                                            MsgBox "Error Bizlink, no se proceso el documento electronico"
                                                                grdDetalle2.MoveFirst
                                                                For ix = 1 To RsDetalle2.RecordCount
                                                                    Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle2.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                                                    Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                    Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle2.Columns(1).Text & "'"
                                                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                    RsDetalle2.MoveNext
                                                                Next ix
                                                                xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                                xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                                Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                                'Sleep 1000
                                                                
                                                                Exit Sub
                                                         End If
                                                    
                                                     
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                             Else
                                                                 cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                             End If
                                                         End If
                                                     End If
                                                End If
                                            ElseIf lFEGesa Then
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                        MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                                        'Exit Sub
                                                    End If
                                                    If lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                                    Else
                                                        If lImpresionCodigoBarras Then
                                                            Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                        Else
                                                            cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                        End If
                                                    End If
                                                End If
                                            Else 'INFOFACT
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                        'Exit Sub
                                                    End If
'                                                    If RsImpresion!Ruc <> "" Then
'                                                        If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                            Exit Sub
'                                                        End If
'                                                    End If
                                                    'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                    If tCodigoFE = "000" Then
                                                        If lQRFE Then
                                                            Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                        Else
                                                            If lImpresionCodigoBarras Then
                                                                Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                            Else
                                                                cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                       End If
                                   End If
                                   '---------------------------------------
                            
                            
                                 'Configura la Impresora
                                 Imprimir (sImp)
                                 sPedido = ""
                                 Printer.FontName = sFont
                                 Printer.FontBold = False
                                     
                                If pais = "003" And Calcular("select isnull(tpuertoserial,'') as codigo from timpresora where tcaja='" & sCaja & "' and timpresora='" & sImp & "'", Cn) <> "" Then  ' Pais Argentina
                                    If ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion) = 0 Then
                                        RevierteDocumento RsDetalle2 ', "", "", sCaja
                                        MsgBox "No se pudo Procesar los documentos, favor de verificar los Datos!!", vbInformation, sMensaje
                                        Screen.MousePointer = vbDefault
                                        RsDetalle2.Requery
                                        Exit Sub
                                    End If
                                Else
                                 
                                 'IMPRESION
                                 If wConsumo Then
                                       If RsTipoDocumento!tFormulario = "01" Then
                                          If RsTipoDocumento!Cliente Then
                                          
                                               'FACTURACION ELECTRONICA
                                               If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                    ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                               Else
                                                   ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                               End If
                                                   NFactura = sCorrela
                                          Else
                                          
                                               'FACTURACION ELECTRONICA
                                               If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                   ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                               Else
                                                   ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                               End If
                                             
                                          End If
                                          
                                       ElseIf RsTipoDocumento!tFormulario = "02" Then
                                          If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                                          
                                               If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                                   ImprimeFormatoAConsumo sDocumento
                                                   Kill App.path & "\fact.bmp"
                                               Else
                                                   ImprimeFormatoAConsumo sDocumento
                                               End If
                                          ElseIf lFacturacionE And lFECarbajal Then
                                                ImprimeFormatoAConsumo sDocumento
                                                
                                                Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                If fso1.FileExists(App.path & "\fact.bmp") Then
                                                    Kill App.path & "\fact.bmp"
                                                End If
                                          Else
                                          
                                               If RsTipoDocumento!Cliente Then
                                                  ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                                  NFactura = sCorrela
                                               Else
                                                  ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                               End If
                                          End If
                                          
                                       Else
                                          If RsTipoDocumento!Cliente Then
                                             ImprimeFacturaVariableConsumo RsImpresion, sDetalleConsumo, sEmpresa
                                             NFactura = sCorrela
                                          End If
                                       End If
                                    
                                 Else
                                 
                                       If RsTipoDocumento!tFormulario = "01" Then
                                           If RsTipoDocumento!Cliente Then
                                           
                                               'FACTURACION ELECTRONICA
                                               If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                   ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                               Else
                                                   ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                               End If
                                              
                                              NFactura = sCorrela
                                           ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                              ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                           Else
                                               
                                               'FACTURACION ELECTRONICA
                                               If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                   ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                               Else
                                                   ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                               End If
                                              
                                           End If
                                          
                                       ElseIf RsTipoDocumento!tFormulario = "02" Then
                                           If RsTipoDocumento!Cliente Then
                                              ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                              NFactura = sCorrela
                                           ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                              ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                           Else
                                              ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                           End If
                                       Else
                                           If RsTipoDocumento!Cliente Then
                                              ImprimeFacturaVariable RsImpresion, sEmpresa
                                              NFactura = sCorrela
                                           ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                              ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                           Else
                                              'ImprimeBoletaN RsImpresion, sEmpresa
                                           End If
                                       End If
                                    
                                 End If
                               End If ' FIN ARGENTINA
                                                
                            End If

                            If pais = "002" And lFEEcuador = False Then
                               sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                               GeneraFacturaElectronica sXML, sDocumento
                            End If
                            
                            If lFEEcuador Then
                             If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                                 MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                             End If
                            End If

                            
               End If
               
                '-----------------------
                If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                     If Not FacturarTCPIP(3, sDocumento, 0) Then
                        MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                     End If
                End If
                '------------------------
               rstTiposPedido.MoveNext
                                               
         Next P
    End If
     
    'Actualiza la Cuenta Corriente
    sTotal = Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'X'", Cn)
    
    Cn.Execute "update TDELIVERY set nConsumo = nConsumo - " & sTotal & ", lReplica=1 where TCODIGODELIVERY ='" & sCtaCte & "'"
    
    'Borra lo facturado
    Cn.Execute "delete from " & sTemporal & " where Estado = 'X'"
    RsDetalle2.Requery
    CalculaMonto
        
    'Borra el temporal
    Cn.Execute "drop table " & sRapido
            
    Set RsSuma = Nothing
    frmCtaCte.RsCabecera.Requery
                    
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdTipoDocumento2_Click(Index As Integer)
On Error GoTo fin:
    Dim RsSuma As Recordset
    Dim RsGuiaFactura As Recordset
    Dim nRespuesta As Integer
    Dim RsItems As Recordset
    Dim sRapido As String
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim sTotal As Double
    Dim xDescuento As Double
    Dim xVenta As Double
    Dim xTVenta As Double
    Dim xImpuesto1 As Double
    Dim xImpuesto2 As Double
    Dim xImpuesto3 As Double
    Dim xNeto As Double

    'FACTURACION_E_PERU
    
    
    lImprimeAlternativa = False
    tAutorizacion = ""
    tCodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    'FACTURACION OFISIS
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    
    Dim oComandoCabeceraOfisis1 As clsComando
    Dim oComandoDetalleOfisis1 As clsComando
    Dim oComandoFirmaDocumentoOfisis1 As clsComando
    
    '------VALIDA CORREO----------
    Dim sTipoDocum As String
    Dim lValidaEmail As Boolean
    Dim sEmail As String
    '-----------------------------
    
    Dim rdi As Integer
    Dim RsCantDocumentos As Recordset
    Dim ix As Integer
    Dim xUltimoCorrelativo As String
    Dim RsTotalDocumento As Recordset
    
    'En blanco Grilla por facturar
    If RsDetalle6.RecordCount = 0 Then
       Exit Sub
    End If
        
    Dim RsCorrela As Recordset
                
   'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento2(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    End If
          
    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
    If nMontoPedidoFacturar < CDbl(Label11.Caption) Then
        MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
        wEnter = False
        Exit Sub
    End If
    End If

    'Consistencia Cortesia
    Dim sCortesia As String
    Dim sUsuarioAutoriza As String
    
    sCortesia = ""
    sUsuarioAutoriza = ""
    
    If RsTipoDocumento!TTipoEmision = "00" Then
       If Supervisor("04") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       
       sUsuarioAutoriza = sVar1
       sTemp = ""
       
       Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
       
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
       frmBusquedaRapida.nPredeterm = 1
       frmBusquedaRapida.Show vbModal
       
       If wEnter = True Then
          sCortesia = sCodigo
       Else
          Exit Sub
       End If
       
       
    Else
       If Supervisor("13") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       sUsuarioAutoriza = sVar1
    End If
        
    'validacortesia
    If RsTipoDocumento!TTipoEmision = "00" And sCortesia <> "" Then
        Dim nTotalCortesiaActual As Double
        Dim nTopeCortesia As Double
        Dim nTotalDocActual As Double
        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
        If nTopeCortesia > 0 Then
'                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and month(fregistro)=month(getdate()) ", Cn)
                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and YEAR(FREGISTRO)=YEAR(GETDATE()) AND month(fregistro)=month(getdate()) ", Cn)
                nTotalDocActual = val(Label11.Caption)
                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                            Exit Sub
                    End If
        End If
    End If

                
        
    'Proceso de Consumo
    Dim wConsumo As Boolean
    wConsumo = False
    sDetalleConsumo = ""
    
    If RsTipoDocumento!TTipoEmision <> "00" Then
       nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
       If nRespuesta = vbYes Then
          frmKeyBoard.txtResultado = "Por Consumo"
          frmKeyBoard.Show vbModal
          If sDescrip = "" Or Not wEnter Then
             MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
             Exit Sub
          End If
          sDetalleConsumo = sDescrip
          wConsumo = True
       ElseIf nRespuesta = vbCancel Then
          Exit Sub
       End If
    End If
        
    'Consistencia Documento con Cliente
    If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= CDbl(txtTotal2.Caption) Or RsTipoDocumento!Monto = 0) And sCliente = "" Then
       MsgBox "Error: El Documento no tiene Cliente", vbCritical, sMensaje
       Exit Sub
         'imprimedni
    Else
        'imprimedni
        Dim rstl As ADODB.Recordset
        Set rstl = New ADODB.Recordset
        Set rstl = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & RsTipoDocumento!TTipoEmision & "','" & sCliente & "'", Cn)
        If Not (rstl.EOF Or rstl.BOF) Then
            rstl.MoveFirst
            If rstl.Fields(0) <> "ok" Then
                MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                Exit Sub
            End If
        End If
    End If
                                                                                            
    If Not RsTipoDocumento!Cliente Then
       sCliente = ""
    End If
    
    TimpresionDolaresDelivery = False
        '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCtaCte & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    '------VALIDA CORREO -----------------
    sTipoDocum = RsTipoDocumento!TTipoEmision
    lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
    
    If lValidaEmail = True And sCliente <> "" Then
       sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
    
       If sEmail = "" Then
          MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
          Exit Sub
       End If
    End If
    '-------------------------------------

    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True And sCliente <> "" Then
        Dim TempUbigeo As String
        Dim TempUrbaniza As String
        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCliente & "' ", Cn)
        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
            MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
            Exit Sub
        End If
    End If
                                                               
    'Crea el temporal para el detalle
    Screen.MousePointer = vbHourglass
    sRapido = dbTemporal(sCaja, 19, "tDocumento", "nVarChar(20)", _
                                    "tItem", "nVarChar(3)", _
                                    "tCodigoPedido", "nVarChar(10)", _
                                    "tCodigoProducto", "nVarChar(7)", _
                                    "nPrecioNeto", "Float", _
                                    "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                    "nPrecioVenta", "Float", _
                                    "nRecargo", "Float", _
                                    "nDescuento", "Float", _
                                    "nCantidad", "Float", _
                                    "nPrecioOficial", "Float", _
                                    "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                    "nVenta", "Float", "nItem", "nVarChar(3)", "ttipopedido", "nvarchar(2)")
                                                                                                                                                                                                                    
                                                                                                                                                                                                                    
    'Inserta Detalle de Documento en un temporal
    
    Set RsGuiaFactura = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tGuiaTransporte in (select tGuiaTransporte from " & sTemporal & " where estado='A')", Cn)
    If RsGuiaFactura.EOF Then
       MsgBox "Error Fatal: No se encontraron Guias a Facturar", vbCritical, sMensaje
       Exit Sub
    End If
    
    For i = 1 To RsGuiaFactura.RecordCount
        Isql = "insert into " & sRapido & _
               "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem,ttipopedido ) " & _
               "select ' ' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tItem, ttipopedido From DPEDIDO where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "' and (isnull(tFacturado,'0')='0' Or len(ltrim(tFacturado))=0) and tEstadoItem ='N'"
        Cn.Execute Isql
        Cn.Execute "update MPEDIDO set tEstadoPedido = '02', tOrdenCompra = '" & txtOrdenCompra2.Text & "'  where tCodigoPedido ='" & RsGuiaFactura!tCodigoPedido & "'" 'FERM - AGREGAR tOrdenCompra
        RsGuiaFactura.MoveNext
    Next i
        
    'Diferentes tipos de pedido
    Dim tipPedido As String
    Dim P As Integer
    Dim rstTiposPedido As New ADODB.Recordset
        
    'Distintos tipos de pedidos
    Set rstTiposPedido = Lib.OpenRecordset("select ttipopedido from " & sRapido & " group by ttipopedido", Cn)
    
    If Not (rstTiposPedido.EOF Or rstTiposPedido.BOF) Then
        rstTiposPedido.MoveFirst
         For P = 1 To rstTiposPedido.RecordCount
                RsTipoDocumento.Requery
                RsTipoDocumento.MoveFirst
                RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento2(Index).Caption & "'")
                tipPedido = rstTiposPedido.Fields("ttipopedido")
                Set RsCorrela = Lib.OpenRecordset("select * from " & sRapido & " where ttipopedido='" & tipPedido & "'", Cn)
                RsCorrela.MoveFirst
                For i = 1 To RsCorrela.RecordCount
                    RsCorrela!tItem = Mid("000", 1, 3 - Len(Trim(str(i)))) & Trim(str(i))
                    RsCorrela.MoveNext
                Next i
                                            
                 Select Case pais
                     Case "001" 'Bolivia
                        tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                        tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                        If tAutorizacion <> "" And tDosificacion <> "" Then
                           Else
                            MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                            Exit Sub
                        End If
                     Case "002"
                           tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
                     Case Else 'Peru, Ecuador
                        tAutorizacion = ""
                        tCodigoControl = ""
                        tDosificacion = ""
                        
                 End Select
                                                     
                 'Verifica y genera la cantidad de documentos
                 If RsTipoDocumento!tFormulario <> "01" And Not wConsumo Then
                        Dim xDocumento(50) As String
                        Dim nTotal As Integer
                        nTotal = 1
                        
                        'Factura
                        'Genera y Actualiza los Numero de Documento
                        sSerie = RsTipoDocumento!tSerie
                        sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                        sPrefijo = RsTipoDocumento!prefijo
                        sTipoDocumento = RsTipoDocumento!TTipoEmision
                        sImp = RsTipoDocumento!timpresora
                        sDocumento = sPrefijo & sSerie & sCorrela
                        sResumen = RsTipoDocumento!lResumen
                        xDocumento(nTotal) = sDocumento
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        'Inserta Detalle de Documento en un temporal
                        grdDetalle6.MoveFirst
                        For i = 1 To RsDetalle6.RecordCount
                           Cn.Execute "update DPEDIDO set tFacturado = 'F', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & grdDetalle6.Columns(0).Text & "' and (Len(lTrim(tFacturado)) = 0 Or IsNull(tFacturado,'0')='0') and tEstadoItem ='N'"
                           Cn.Execute "update DPEDIDO set tFacturado = 'C', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & grdDetalle6.Columns(0).Text & "' and tEstadoItem ='N' and substring(tDocumento,1,1)='0'"
                           Cn.Execute "update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & grdDetalle6.Columns(0).Text & "'"
                           Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='02', tDocumento='" & sDocumento & "' Where tGuiaTransporte='" & grdDetalle6.Columns(0).Text & "'"
                           grdDetalle6.MoveNext
                        Next i
                    
                       'PARA NO FISCALES
                       Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & grdDetalle6.Columns(0).Text & "' and tDocumento ='" & sDocumento & "'"
                       
                        
                        Set RsRapido = Lib.OpenRecordset("select * from " & sRapido & " where ttipopedido='" & tipPedido & "'", Cn)
                        RsRapido.MoveFirst
                        i = 1
                        sTotal = 0
                                                      
                        Do While Not RsRapido.EOF
                              sTotal = sTotal + 1
                              If i >= nFItem + 1 Then
                                 Isql = "Insert into MDOCUMENTO " & _
                                        "     ( tDocumento, tTipoDocumento, tCortesia, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro, fdiaContable, lImpresionMonedaExtranjera) " & _
                                        "Values(   '" & sDocumento & "', " _
                                                & "'" & sTipoDocumento & "', " _
                                                & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                                & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                                & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                                & "'" & sCaja & "', " _
                                                & "'" & sSalon & "', " _
                                                & "'" & sTurno & "', " _
                                                & "'" & sUsuario & "', " _
                                                & "'" & sUsuarioAutoriza & "', " _
                                                & "'" & IIf(wConsumo, "C", "D") & "', " _
                                                & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                                & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " ) "
                                 Cn.Execute Isql
                                 
                                 'Calcula el total de la cabecera
                                 Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento),0) as nDescuento " & _
                                                                " from " & sRapido & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
                                                                
                                 xDescuento = RsSuma!nDescuento
                                 xVenta = RsSuma!nVenta
                                 xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                                 xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                                 xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                                 xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                                                                
                                Select Case pais
                                    Case "001"
                                        tCodigoControl = devuelveCodigoControl(sCaja, Mid(sDocumento, 7, 9), tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                                End Select
                                                                
                                 'Actualiza el Documento con el Temporal
                                 Isql = "Update MDOCUMENTO set nNeto= " & xNeto & ", " & _
                                                              "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                              "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                              "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                              "nRecargo = 0, " & _
                                                              "nDescuento = " & xDescuento & ", " & _
                                                              "nVenta = " & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " & _
                                                              "nPrecioOficial = " & RsSuma!nVenta & ", " & _
                                                              "tautorizacion = '" & tAutorizacion & "' , " & _
                                                              "tcodigocontrol = '" & tCodigoControl & "' , " & _
                                                              " lreplica=1  where tDocumento = '" & sDocumento & "'"
                                 Cn.Execute Isql
                                 
                                 
                                 'Inserta todos los items del temporal en el DETALLE
                                 Isql = "Insert into DDOCUMENTO " & _
                                        "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                        "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                                        "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                        "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sRapido & " where tDocumento ='" & sDocumento & "'"
                                 Cn.Execute Isql
                                 
                                 
                                   'SACA LA VUELTA 06082021
                                  'aqui 14102021 GCAA

                                   Dim oComandox As clsComando
                                   Set oComandox = New clsComando

                                   Set oComandox = New clsComando
                                    If Not oComandox.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                    End If

                                    oComandox.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento

                                    If Not oComandox.GetParamOK Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                    End If

                                    If Not oComandox.ExecSP Then
                                       Set oComandox = Nothing
                                       Exit Sub
                                     End If
                                   'fin 14102021 GCAA
                                 
                                 
                                 'Actualiza Base de Datos Detalle del Pedido
                                 'Genera y Actualiza los Numero de Documento
                                 RsTipoDocumento.Requery
                                 RsTipoDocumento.MoveFirst
                                 RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento2(Index).Caption & "'")
                                 If RsTipoDocumento.EOF Then
                                    MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
                                    Exit Sub
                                 End If
                                         
                                 sSerie = RsTipoDocumento!tSerie
                                 sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                                 sPrefijo = RsTipoDocumento!prefijo
                                 sTipoDocumento = RsTipoDocumento!TTipoEmision
                                                                  
                                 sImp = RsTipoDocumento!timpresora
                                 sDocumento = sPrefijo & sSerie & sCorrela
                                                                 
                                 sResumen = RsTipoDocumento!lResumen
                                 nTotal = nTotal + 1
                                 
                                 xDocumento(nTotal) = sDocumento
                                 Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                 
                                 i = 2
                                 
                                 RsRapido!tDocumento = sDocumento
                                 RsRapido.Update
                                 
                                 sPedido = IIf(IsNull(RsRapido!tCodigoPedido), " ", RsRapido!tCodigoPedido)
                                 
                              Else
                                 RsRapido!tDocumento = sDocumento
                                 RsRapido.Update
                                 sPedido = IIf(IsNull(RsRapido!tCodigoPedido), " ", RsRapido!tCodigoPedido)
                                 i = i + 1
                              End If
                              
                              RsRapido.MoveNext
                              
                        Loop
                            Isql = "Insert into MDOCUMENTO " & _
                                   "     ( tDocumento, tTipoDocumento, tCortesia, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro, fdiacontable, lImpresionMonedaExtranjera) " & _
                                   "Values(   '" & sDocumento & "', " _
                                           & "'" & sTipoDocumento & "', " _
                                           & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                           & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                           & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                           & "'" & sCaja & "', " _
                                           & "'" & sSalon & "', " _
                                           & "'" & sTurno & "', " _
                                           & "'" & sUsuario & "', " _
                                           & "'" & sUsuarioAutoriza & "', " _
                                           & "'" & IIf(wConsumo, "C", "D") & "', " _
                                           & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                           & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & ") "
                                 
                           Cn.Execute Isql
                           
                           'Calcula el total de la cabecera
                           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento),0) as nDescuento " & _
                                                          " from " & sRapido & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
                                                                                                         
                           xVenta = RsSuma!nVenta
                           xDescuento = RsSuma!nDescuento
                           
                           xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                           xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                           
                           xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                           xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                           
                            Select Case pais
                               Case "001"
                                    tCodigoControl = devuelveCodigoControl(sCaja, Mid(sDocumento, 7, 9), tAutorizacion, tDosificacion, sCliente, xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3)
                            End Select
                                                  
                           'Actualiza el Documento con el Temporal
                           Isql = "Update MDOCUMENTO set nNeto= " & xNeto & ", " & _
                                                        "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                        "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                        "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                        "nRecargo = 0, " & _
                                                        "nDescuento = " & xDescuento & ", " & _
                                                        "tautorizacion='" & tAutorizacion & "', " & _
                                                        "tcodigocontrol='" & tCodigoControl & "'," & _
                                                        "nVenta = " & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " & _
                                                        "nPrecioOficial = " & RsSuma!nVenta & " " & _
                                                        " ,lreplica=1  where tDocumento = '" & sDocumento & "'"
                           Cn.Execute Isql
                                                                  
                           'Inserta todos los items del temporal en el DETALLE
                           Isql = "Insert into DDOCUMENTO " & _
                                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                                  "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sRapido & " where tDocumento ='" & sDocumento & "'"
                           Cn.Execute Isql
                           
                             'SACA LA VUELTA 06082021
                                  'aqui 14102021 GCAA

                                   Dim oComandoxx As clsComando
                                   Set oComandoxx = New clsComando

                                   Set oComandoxx = New clsComando
                                    If Not oComandoxx.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                    End If

                                    oComandoxx.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento

                                    If Not oComandoxx.GetParamOK Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                    End If

                                    If Not oComandoxx.ExecSP Then
                                       Set oComandoxx = Nothing
                                       Exit Sub
                                     End If
                                   'fin 14102021 GCAA
                           
                           
                            If lPagoAntesImpresion Then
                                For i = 1 To nTotal
                                    wEnter = True
                                    
                                    Set RsTotalDocumento = Lib.OpenRecordset("select nventa from MDOCUMENTO where tDocumento ='" & xDocumento(i) & "'", Cn)
                                    nCargo = Round(RsTotalDocumento!nVenta, 2)
                                    sDocumento = xDocumento(i)
                                    sFormulario = ""
                                    sTipo = ""
                                    frmPago.Show vbModal
                                    
                                    If wEnter = False Then
                                        grdDetalle6.MoveFirst
                                        For ix = 1 To RsGuiaFactura.RecordCount
                                            Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle6.Columns(1).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & grdDetalle6.Columns(1).Text & "'"
                                            Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & grdDetalle6.Columns(1).Text & "'"
                                            For rdi = 0 To RsCantDocumentos.RecordCount - 1
                                                Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                                                RsCantDocumentos.MoveNext
                                            Next rdi
                                            RsGuiaFactura.MoveNext
                                        Next ix
                            
                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Exit Sub
                                    End If
                                Next i
                            End If
                           
                           
                           
                           
                           'Actualiza Base de Datos Detalle del Pedido
                           Isql = "UPDATE DPEDIDO SET DPEDIDO.tFacturado = 'F', DPEDIDO.tDocumento = " & sRapido & ".tDocumento " & _
                                  "FROM  DPEDIDO INNER JOIN " & sRapido & " ON DPEDIDO.tCodigoPedido = " & sRapido & ".tCodigoPedido AND DPEDIDO.tItem = " & sRapido & ".nItem "
                           Cn.Execute Isql
                           
                           Isql = "UPDATE DPEDIDO SET DPEDIDO.tFacturado = 'C', DPEDIDO.tDocumento = " & sRapido & ".tDocumento " & _
                                  "FROM DPEDIDO INNER JOIN " & sRapido & " ON DPEDIDO.tCodigoPedido = " & sRapido & ".tCodigoPedido AND substring(DPEDIDO.tDocumento,1,1)='0'"
                           Cn.Execute Isql
                            
                           
                           'PARA NO FISCALES
                           Dim RST As New ADODB.Recordset
                           Set RST = Lib.OpenRecordset("select tdocumento from " & sRapido & " order by tdocumento", Cn)
                           If Not (RST.EOF Or RST.BOF) Then
                                RST.MoveFirst
                                Do While Not RST.EOF
                                    
                                    Isql = "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tDocumento ='" & RST!tDocumento & "'"
                                    Cn.Execute Isql
                                    RST.MoveNext
                                Loop
                           End If
                            
                            
                           For i = 1 To nTotal
                               'Imprime Documentos
                               If wConsumo = False And lDescripcionAlternativa = True Then
                                  If validaImpresionAlternativa(sDocumento) = False Then
                                             If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                                     lImprimeAlternativa = True
                                             End If
                                  End If
                               End If
                                 '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                                 ' ELDCQ 15/11/2017
                                 If lImprimeAlternativa = False Then
                                    If lDocumentoAgrupado Then
                                        Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',1"
                                        'FACTURACION_E_PERU
                                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',2"
                                    Else
                                        Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',3"
                                        'FACTURACION_E_PERU
                                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',4"
                                    End If
                                Else
                                    If lDocumentoAgrupado Then
                                        Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',5"
                                        'FACTURACION_E_PERU
                                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',6"
                                    Else
                                        Isql = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',7"
                                        'FACTURACION_E_PERU
                                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & xDocumento(i) & "',8"
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------------
                                'FACTURACION_E_PERU
                                'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                                xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & xDocumento(i) & "')", Cn)
                                xImpresioDE = Mid(sDocumento, 1, 1)
                                '---------------------------------------
                          
                                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                             
                               If RsImpresion.RecordCount = 0 Then
                                  LimpiaRs
                                  MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                               Else
                                'SUNAT
                                  numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                                  codigoImpresora = sImp
                                'SUNAT
                                  Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & xDocumento(i) & "' "
    
                                    'impresion imagen
                                Set rstFuente = New ADODB.Recordset
                                imageCab.Picture = Nothing
                                imagepIE.Picture = Nothing
                                Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                                imageCab.DataField = "foto"
                                Set imageCab.DataSource = rstFuente
                                imagepIE.DataField = "fotoPie"
                                Set imagepIE.DataSource = rstFuente
                                
                                
                               'FACTURACION_E_PERU
                               If pais = "000" Then
                                   If lFacturacionE Then
                                       If lFEOfisis Then 'OFISIS
                                                
                                                 '----CABECERA
                                                Set oComandoCabeceraOfisis = New clsComando
                                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, xDocumento(i)
                    
                                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoCabeceraOfisis.ExecSP Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                '----FIRMA DOCUMENTO OFISIS
                                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, xDocumento(i)
                        
                                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    
                                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                                    fDocumento = Mid(xDocumento(i), 1, 1) + Mid(xDocumento(i), 4, 3) + "-" + CStr(CLng(Mid(xDocumento(i), 8, 8)))
                                                    Sleep 3000
                                                    If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                        
                                                    ElseIf lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                    Else
                                                            Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                            If RscadenaCodigoHash.RecordCount > 0 Then
                                                                cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                            End If

                                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                    End If
                                                End If

                                       ElseIf lFESpring Then
                                       
                                       ElseIf lFECarbajal Then
                                       
                                       ElseIf lFEpape Then
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                         Else
                                                             cadenaCodigoHash = PapeMatricial
                                                         End If
                                                     End If
                                                 End If
                                            End If
                                       Else 'INFOFACT
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                If Not INSERTAFE(xDocumento(i), "", 1, "") Then '----CABECERA
                                                    'Exit Sub
                                                End If
'                                                If RsImpresion!Ruc <> "" Then
'                                                    If Not INSERTAFE(xDocumento(i), "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                        Exit Sub
'                                                    End If
'                                                End If
                                                'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                fDocumento = Mid(xDocumento(i), 1, 1) + Mid(xDocumento(i), 4, 3) + Mid(xDocumento(i), 8, 8)
                                                If tCodigoFE = "000" Then
                                                    If lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR(xDocumento(i)))
                                                    Else
                                                        If lImpresionCodigoBarras Then
                                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, xDocumento(i)))
                                                        Else
                                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, xDocumento(i))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                       End If
                                   End If
                            End If
                            '---------------------------------------

                                  'Configura la Impresora
                                  Imprimir (sImp)
                                  Printer.FontName = sFont
                                  Printer.FontBold = False
                                                
                                   If wConsumo Then
                                      If RsTipoDocumento!tFormulario = "01" Then
                                         If RsTipoDocumento!Cliente Then
                                            ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                            NFactura = sCorrela
                                         Else
                                            ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                         End If
                                      Else
                                         If RsTipoDocumento!Cliente Then
                                            ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                            NFactura = sCorrela
                                         Else
                                            ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                         End If
                                      End If
                                      
                                   Else
                                       If RsTipoDocumento!tFormulario = "01" Then
                                       
                                              If RsTipoDocumento!Cliente Then
                                                 
                                                    If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                       ImprimeFacturaElectronica RsImpresion, Nothing, "", Nothing, Nothing, "", TimpresionDolaresDelivery
                                                    Else
                                                       ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                                    End If
                                                
                                                    NFactura = sCorrela
                                                 
                                              ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                                    If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                                       ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                                    End If
                                                 
                                              Else
        
                                                     If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                        ImprimeBoletaElectronica RsImpresion, Nothing, "", Nothing, Nothing, "", TimpresionDolaresDelivery
                                                     Else
                                                        ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                                     End If
                                                     
                                              End If
                                       Else
                                              If lFacturacionE Then
                                                        'FACTURACION_E_PERU
                                                        'FORMATO A4
                                                        If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                                           ImprimeFormatoA xDocumento(i)
                                                           Kill App.path & "\fact.bmp"
                                                        Else
                                                           ImprimeFormatoA xDocumento(i)
                                                        End If
                                              Else
                                                        If RsTipoDocumento!Cliente Then
                                                           ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                                           NFactura = sCorrela
                                                        ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                                           If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                                              ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                                           End If
                                                        Else
                                                           ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                                        End If
                                              End If
                                       End If
                                       
                                   End If
                                  LimpiaRs
                               End If

'                               If pais = "002" Then
'                                  sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                                  GeneraFacturaElectronica sXML, sDocumento
'                               End If
                               
                                If pais = "002" And lFEEcuador = False Then
                                   sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                                   GeneraFacturaElectronica sXML, sDocumento
                                End If
                                
                                If lFEEcuador Then
                                 If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                                     MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                                 End If
                                End If

                           Next i

                           
                '------------- FORMATO TICKET --------------------
                Else
                           'Genera y Actualiza los Numero de Documento
                           sSerie = RsTipoDocumento!tSerie
                           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                           sPrefijo = RsTipoDocumento!prefijo
                           sTipoDocumento = RsTipoDocumento!TTipoEmision
                           sImp = RsTipoDocumento!timpresora
                           sDocumento = sPrefijo & sSerie & sCorrela
                           sResumen = RsTipoDocumento!lResumen
                           'Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                  
                           'Inserta Detalle de Documento en un temporal
                           RsGuiaFactura.MoveFirst
                           For i = 1 To RsGuiaFactura.RecordCount
                               Cn.Execute "update DPEDIDO set tFacturado = 'F', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "' and (Len(lTrim(tFacturado)) = 0 Or IsNull(tFacturado,'0')='0') and tEstadoItem ='N'"
                               Cn.Execute "update DPEDIDO set tFacturado = 'C', tDocumento ='" & sDocumento & "' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "' and tEstadoItem ='N' and substring(tDocumento,1,1)='0'"
                               Cn.Execute "update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & RsGuiaFactura!tCodigoPedido & "'"
                               
                               RsGuiaFactura.MoveNext
                           Next i
                           
                           RsDetalle6.MoveFirst
                           For i = 1 To RsDetalle6.RecordCount
                               Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='02', tDocumento='" & sDocumento & "', tOrdenCompra = '" & txtOrdenCompra2.Text & "' Where tGuiaTransporte='" & RsDetalle6!tGuiaTransporte & "'" 'FERM - Agregar tOrdenCompra
                               RsDetalle6.MoveNext
                           Next i
                           'PARA NO FISCALES
                           Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & grdDetalle6.Columns(0).Text & "' and tDocumento ='" & sDocumento & "'"
                           
                           'Inserta todos los items del temporal en el DETALLE
                            Isql = "Insert into DDOCUMENTO " & _
                                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                                  "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sRapido & " where ttipopedido='" & tipPedido & "'"
                           Cn.Execute Isql
                           
                           'Calcula el total de la cabecera
                           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento),0) as nDescuento " & _
                                                          " from " & sRapido & " where ttipopedido='" & tipPedido & "'  group by tDocumento", Cn)
                                                           
                           xDescuento = RsSuma!nDescuento
                           xVenta = RsSuma!nVenta
                           xImpuesto1 = IIf(RsSuma!nImpuesto1 = 0, 0, xVenta * RsSuma!nImpuesto1 / RsSuma!nVenta)
                           xImpuesto2 = IIf(RsSuma!nImpuesto2 = 0, 0, xVenta * RsSuma!nImpuesto2 / RsSuma!nVenta)
                           xImpuesto3 = IIf(RsSuma!nImpuesto3 = 0, 0, xVenta * RsSuma!nImpuesto3 / RsSuma!nVenta)
                           xNeto = xVenta - xImpuesto1 - xImpuesto2 - xImpuesto3
                                                         
                           Select Case pais
                                Case "001"
                                    tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                           End Select
                           
                           'Inserta el Documento
                           ' Ferm - Agregar orden de compra
                           Isql = "Insert into MDOCUMENTO " & _
                                  "     ( tDocumento, tTipoDocumento, tCortesia, tcodigoCliente, tEstadoDocumento, tCaja, tTurno, nNeto, nDescuento, nRecargo, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, tSalon, tUsuario, tUsuarioAutoriza, tEmision, tConsumo, fRegistro,tautorizacion,tcodigocontrol, fDiaContable, lImpresionMonedaExtranjera, tOrdenCompra ) " & _
                                  "Values(   '" & sDocumento & "', " _
                                          & "'" & sTipoDocumento & "', " _
                                          & "'" & IIf(Mid(sDocumento, 1, 1) = "0", sCortesia, "") & "', " _
                                          & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                                          & "'" & IIf(Mid(sDocumento, 1, 1) = "0", "02", "01") & "', " _
                                          & "'" & sCaja & "', " _
                                          & "'" & sTurno & "', " _
                                          & xNeto & ", " _
                                          & xDescuento & ", " _
                                          & "0 , " _
                                          & RsSuma!nVenta & ", " _
                                          & xImpuesto1 & ", " _
                                          & xImpuesto2 & ", " _
                                          & xImpuesto3 & ", " _
                                          & xNeto + xImpuesto1 + xImpuesto2 + xImpuesto3 & ", " _
                                          & "'" & sSalon & "', " _
                                          & "'" & sUsuario & "', " _
                                          & "'" & sUsuarioAutoriza & "', " _
                                          & "'" & IIf(wConsumo, "C", "D") & "', " _
                                          & "'" & IIf(wConsumo, sDetalleConsumo, "") & "', " _
                                          & " getdate(),'" & tAutorizacion & "','" & tCodigoControl & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & ", '" & Replace(txtOrdenCompra2.Text, "'", "''") & "')"
                                          
                           Cn.Execute Isql
                           
                            
                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                           
                            If lPagoAntesImpresion Then
                                   Set RsTotalDocumento = Lib.OpenRecordset("select nventa from MDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                                   nCargo = Round(RsTotalDocumento!nVenta, 2)
                                   sFormulario = ""
                                   sTipo = ""
                                   frmPago.Show vbModal
                                   
                                   If wEnter = False Then
                                        RsGuiaFactura.MoveFirst
                                        For ix = 1 To RsGuiaFactura.RecordCount
                                        ' FERM - actualizar orden de compra (al ser eliminado/cancelado la operación)
                                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "'"
                                            Cn.Execute "update MPEDIDO set tEstadoPedido = '04', tOrdenCompra = '' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "'"
                                            Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                            
                                             RsGuiaFactura.MoveNext
                                        Next ix
                                         
                                        RsDetalle6.MoveFirst
                                        For i = 1 To RsDetalle6.RecordCount
                                            ' FERM - Actualizar orden de compra (al ser eliminado/cancelado la operación)
                                            Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='01', tDocumento='', tOrdenCompra = '' Where tGuiaTransporte='" & RsDetalle6!tGuiaTransporte & "'"
                                            RsDetalle6.MoveNext
                                        Next i
                                            
                                         xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                         
                                         xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                         
                                         Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                         Exit Sub
                                   Else
                                      Cn.Execute "UPDATE TESTADOINFOFACT SET lEstadoFacturacion = 0 WHERE tDocumento IN (SELECT tGuiaTransporte FROM MGUIATRANSPORTE WHERE tDocumento = '" & sDocumento & "')"
                                   End If
                            End If
                             '-----------------------
                             If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not FacturarTCPIP(2, sDocumento, 0) Then
                                      RsGuiaFactura.MoveFirst
                                        For ix = 1 To RsGuiaFactura.RecordCount
                                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "'"
                                            Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & RsGuiaFactura!tCodigoPedido & "'"
                                            Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                            Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                            
                                             RsGuiaFactura.MoveNext
                                        Next ix
                                         
                                        RsDetalle6.MoveFirst
                                        For i = 1 To RsDetalle6.RecordCount
                                            Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='01', tDocumento='' Where tGuiaTransporte='" & RsDetalle6!tGuiaTransporte & "'"
                                            RsDetalle6.MoveNext
                                        Next i
                                     
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                     
                                    Exit Sub
                                End If
                             End If
                            '------------------------
                                 
                            'Imprime Documentos
                            If wConsumo = False And lDescripcionAlternativa = True Then
                                If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                        lImprimeAlternativa = True
                                End If
                            End If
                            '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                            ' ELDCQ 15/11/2017
                            If lImprimeAlternativa = False Then
                                If lDocumentoAgrupado Then
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
                                Else
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                                End If
                            Else
                                If lDocumentoAgrupado Then
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
                                Else
                                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                                    'FACTURACION_E_PERU
                                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                                End If
                            End If
                            '------------------------------------------------------------------------------------------------------------
                            'FACTURACION_E_PERU
                            'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                            xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                            xImpresioDE = Mid(sDocumento, 1, 1)
                            '---------------------------------------
                                        
                            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                            'impresion imagen
                            Set rstFuente = New ADODB.Recordset
                            imageCab.Picture = Nothing
                            imagepIE.Picture = Nothing
                            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                            imageCab.DataField = "foto"
                            Set imageCab.DataSource = rstFuente
                            imagepIE.DataField = "fotoPie"
                            Set imagepIE.DataSource = rstFuente
                                
                             
                           If RsImpresion.RecordCount = 0 Then
                              LimpiaRs
                              MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                           Else
                              'SUNAT
                              numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                              codigoImpresora = sImp
                              'SUNAT
                              Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "
                               
                               
                               'FACTURACION_E_PERU
                               If pais = "000" Then
                                   If lFacturacionE Then
                                   
                                        If lFEOfisis Then 'OFISIS
                                        
                                                '----CABECERA
                                                Set oComandoCabeceraOfisis1 = New clsComando
                                                If Not oComandoCabeceraOfisis1.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                     Set oComandoCabeceraOfisis1 = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoCabeceraOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                    
                                                If Not oComandoCabeceraOfisis1.GetParamOK Then
                                                     Set oComandoCabeceraOfisis1 = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoCabeceraOfisis1.ExecSP Then
                                                     Set oComandoCabeceraOfisis1 = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                '----FIRMA DOCUMENTO OFISIS
                                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                    Set oComandoFirmaDocumentoOfisis1 = New clsComando
                                                    If Not oComandoFirmaDocumentoOfisis1.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                         Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoFirmaDocumentoOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                                    If Not oComandoFirmaDocumentoOfisis1.GetParamOK Then
                                                         Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoFirmaDocumentoOfisis1.ExecSP Then
                                                         Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                         Exit Sub
                                                    End If
                                                    
                                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                    Sleep 3000
                                                    If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                        
                                                    ElseIf lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                    Else
                                                            Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                            If RscadenaCodigoHash.RecordCount > 0 Then
                                                                cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                            End If

                                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                    End If
                                                End If
                                                
                                       ElseIf lFESpring Then
                                       
                                       ElseIf lFECarbajal Then
                                                Label4.Caption = "   Proceso de envio de documento a InfoFact......."
                                                lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                                lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                    
                                                    FrameFeSpring.Visible = True
                                                    lblPaso1.Visible = True
                                                    lblPaso2.Visible = True
                                                    imgProceso(0).Visible = False
                                                    imgProceso(1).Visible = False
                                                    imgProceso(2).Visible = False
                                                    imgProceso(3).Visible = False
                                                    Sleep 1000
                                                    If Not INSERTAFE_CARVAJAL(sDocumento, "", 0, 0) Then '----CABECERA
                                                            grdDetalle6.MoveFirst
                                                            For ix = 1 To RsDetalle6.RecordCount
                                                                'Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle6.Columns(0).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                                                Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='01', tDocumento='' Where tGuiaTransporte='" & RsDetalle6!tGuiaTransporte & "'"
                                                                Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                                grdDetalle6.MoveNext
                                                            Next ix
                                                            xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                            xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                            imgProceso(2).Visible = True
                                                            imgProceso(3).Visible = True
                                                            Sleep 1000
                                                            FrameFeSpring.Visible = False
                                                            Screen.MousePointer = vbDefault
                                                            MsgBox "Se produjo un incoveniente en el proceso, favor de verificar la información!!", vbInformation
                                                            RsDetalle6.Requery
                                                            Exit Sub
                                                     End If
                                                     imgProceso(0).Visible = True
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     'fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                                             Else
                                                                 cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                                             End If
                                                         End If
                                                     End If
                                                     imgProceso(1).Visible = True
                                                     Sleep 1000
                                                     FrameFeSpring.Visible = False
                                                End If
                                                                                    
                                       ElseIf lFEpape Then
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                         Else
                                                             cadenaCodigoHash = PapeMatricial
                                                         End If
                                                     End If
                                                 End If
                                            End If
                                        ElseIf lFEBiz Then
                                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                
                                                   If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                                           RsDetalle6.MoveFirst
                                                           For ix = 1 To RsDetalle6.RecordCount
                                                               'Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & grdDetalle6.Columns(0).Text & "' and M.tEstadoDocumento <> '04'", Cn)
                                                               Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='01', tDocumento='' Where tGuiaTransporte='" & grdDetalle6.Columns(0).Text & "'"
                                                               Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                               Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                               Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                               
                                                               Cn.Execute "update dpedido set tdocumento='',tfacturado='' Where tguiatransporte= '" & grdDetalle6.Columns(0).Text & "'"
                                                               
                                                               grdDetalle6.MoveNext
                                                           Next ix
                                                           xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                           xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                           Sleep 1000
                                                           Screen.MousePointer = vbDefault
                                                           MsgBox "Se produjo un incoveniente en el proceso, favor de verificar la información!!", vbInformation
                                                           RsDetalle6.Requery
                                                           Exit Sub
                                                    End If
                                                   
                                                     
                                                     'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                     If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                             Else
                                                                 cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                             End If
                                                         End If
                                                     End If
                                                End If
                                        ElseIf lFEGesa Then
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                                End If
                                                If lQRFE Then
                                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                                Else
                                                    If lImpresionCodigoBarras Then
                                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                    Else
                                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                    End If
                                                End If
                                            End If
                                        Else 'INFOFACT
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                    'Exit Sub
                                                End If
'                                                If RsImpresion!Ruc <> "" Then
'                                                    If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                        Exit Sub
'                                                    End If
'                                                End If
                                                'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                If tCodigoFE = "000" Then
                                                    If lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                    Else
                                                        If lImpresionCodigoBarras Then
                                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                        Else
                                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                   End If
                            End If
                            '---------------------------------------
                            'Configura la Impresora
                            Imprimir (sImp)
                            sPedido = ""
                            Printer.FontName = sFont
                            Printer.FontBold = False
                            '--------------------------------------
                              
                            If pais = "003" Then  ' Pais Argentina
                                'Screen.MousePointer = vbHourglass
                                If ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion) = 0 Then
                                    RevierteDocumentoGUIA RsDetalle6 'sDocumento, "", "", sCaja
                                    MsgBox "No se pudo Procesar los documentos, favor de verificar los Datos!!", vbInformation, sMensaje
                                    Screen.MousePointer = vbDefault
                                    RsDetalle6.Requery
                                    Exit Sub
                                End If
                                'Screen.MousePointer = vbDefault
                            Else
                              'IMPRESION
                              If wConsumo Then
                                    If RsTipoDocumento!tFormulario = "01" Then
                                       If RsTipoDocumento!Cliente Then
                                       
                                            'FACTURACION ELECTRONICA
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                            Else
                                                ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                            End If
                                          
                                            NFactura = sCorrela
                                       Else
                                       
                                            'FACTURACION ELECTRONICA
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                            Else
                                                ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                            End If
                                          
                                       End If
                                       
                                    Else
                                       If lFacturacionE Then
                                       
                                                If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                                     ImprimeFormatoAConsumo sDocumento
                                                     Kill App.path & "\fact.bmp"
                                                 Else
                                                     ImprimeFormatoAConsumo sDocumento
                                                 End If
                            
                                       Else
                                    
                                                If RsTipoDocumento!Cliente Then
                                                   ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                                   NFactura = sCorrela
                                                Else
                                                   ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                                End If
                                       
                                       End If
                                       
                                    End If
                              Else
                              
                                    If RsTipoDocumento!tFormulario = "01" Then
                                          If RsTipoDocumento!Cliente Then
                                          
                                             'FACTURACION ELECTRONICA
                                              If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                 ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                              Else
                                                 ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                              End If
                                             
                                             NFactura = sCorrela
                                          ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                             ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                          Else
                                        
                                              'FACTURACION ELECTRONICA
                                              If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                                 ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                              Else
                                                 ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                              End If
                                             
                                          End If
                                    Else
                                          If RsTipoDocumento!Cliente Then
                                               ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                               NFactura = sCorrela
                                          ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                               ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                          Else
                                               ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                          End If
                                    End If
                                 
                              End If
                             End If ' Argentina
                           End If
                        End If
                        
'                        If pais = "002" Then
'                           sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                           GeneraFacturaElectronica sXML, sDocumento
'                        End If
                        If pais = "002" And lFEEcuador = False Then
                           sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                           GeneraFacturaElectronica sXML, sDocumento
                        End If
                        
                        If lFEEcuador Then
                         If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                             MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                         End If
                        End If
                        '-----------------------
                        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                             If Not FacturarTCPIP(3, sDocumento, 0) Then
                                MsgBox ("La confirmacion del Doc. Elec." & sDocumento & " ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                             End If
                        End If
                        '------------------------
                        rstTiposPedido.MoveNext
         Next P
     End If
     
    'Actualiza la Cuenta Corriente
    sTotal = Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'A'", Cn)
    Cn.Execute "update TDELIVERY set nConsumo = nConsumo - " & sTotal & ", lReplica=1 where TCODIGODELIVERY ='" & sCtaCte & "'"

    'Borra lo facturado
    Cn.Execute "delete from " & sTemporal & " where Estado = 'A'"
    RsDetalle6.Requery
    CalculaMonto
        
    'Borra el temporal
    Cn.Execute "drop table " & sRapido
            
    Set RsSuma = Nothing
    frmCtaCte.RsCabecera.Requery
                    
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox Error
    Exit Sub
End Sub



Private Sub ImprimeFormatoAConsumo(ByVal nDocumento As String)

                        Dim ReporteC As New dsrBoletaC
                        
                        If RsTipoDocumento!lImprimeImageCab Then
                           iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                        End If
                        
                        ReporteC.DiscardSavedData
                        ReporteC.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                        
                        If xImpresioDE = "B" Then
                           ReporteC.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                        ElseIf xImpresioDE = "F" Then
                           ReporteC.Text13.SetText "FACTURA ELECTRONICA"
                        End If
                                                            
                        ReporteC.Text8.SetText sRazonSocial
                        ReporteC.ReportTitle = sDireccion
                        ReporteC.Text15.SetText sTelefono
                        ReporteC.Text33.SetText sFax
                        ReporteC.Text16.SetText sRUC
                        ReporteC.Text50.SetText sWeb
                        
                        ReporteC.Text31.SetText sDetalleConsumo
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                           ReporteC.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                        ReporteC.Text4.SetText xMontoTexto
                        ReporteC.Text32.SetText xImpresionFE

                        If lFacturacionE And lFECarbajal Then
                            If lQRFE = False And lImpresionCodigoBarras = False Then
                                'ReporteC.ImagenHash.Suppress
                                ReporteC.CodigoHash.SetText cadenaCodigoHash
                            Else
                                'ReporteC.CodigoHash.Suppress
                            End If
                        End If


'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = ReporteC
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                        
                        ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ReporteC.PaperOrientation = crPortrait
                        ReporteC.PrintOut False, 1, False, 1, 1
                        '----------------
                                                          
                    
                        If iImagenCab Then
                           Kill App.path & "\cliente.jpg"
                        End If
End Sub


Private Sub ImprimeFormatoA(ByVal nDocumento As String)
                    Dim Reporte As New dsrBoleta

                    If RsTipoDocumento!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                
                    Reporte.DiscardSavedData
                    Reporte.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                    
                    If xImpresioDE = "B" Then
                       Reporte.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       Reporte.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                                                        
                    Reporte.Text8.SetText sRazonSocial
                    Reporte.ReportTitle = sDireccion
                    Reporte.Text15.SetText sTelefono
                    Reporte.Text14.SetText sFax
                    Reporte.Text16.SetText sRUC
                    Reporte.Text50.SetText sWeb
                    
                    'Miguel rivera---------------
                    Reporte.Text43.SetText sMonN
                    Reporte.Text44.SetText sMonN
                    Reporte.Text45.SetText sMonN
                    Reporte.Text46.SetText sMonN
                    Reporte.Text47.SetText sMonN
                    Reporte.Text48.SetText sMonN
                    '-----------------------------
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                    Reporte.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                    Reporte.Text4.SetText xMontoTexto
                    Reporte.Text31.SetText xImpresionFE

                    If lFacturacionE And lFECarbajal Then
                        If lQRFE = False And lImpresionCodigoBarras = False Then
                            'Reporte.ImagenHash.Suppress
                            Reporte.CodigoHash.SetText cadenaCodigoHash
                        Else
                            'Reporte.CodigoHash.Suppress
                        End If
                    End If

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                    
                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    Reporte.PrintOut False, 1, False, 1, 1
                    '----------------
                    
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub

Private Sub cmdTipoGuia_Click(Index As Integer)
    Dim RsGuiaGrupoProducto As Recordset
    Dim RsCorrela As Recordset
    Dim sGuiaTemporal As String
    Dim RsGuia As Recordset
    Dim RsCantidad As Recordset
    Dim clasexml As New clsxml
    Dim Ruta As String
    Dim sTipoDocumento As String
    'En blanco Grilla por facturar
    If RsDetalle4.RecordCount = 0 Then
       Exit Sub
    End If
        
    'Chequea Consistencia
    RsTipoGuia.Requery
    RsTipoGuia.MoveFirst
    RsTipoGuia.Find ("Descripcion='" & cmdTipoGuia(Index).Caption & "'")
    If RsTipoGuia.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    Else
        sTipoDocumento = RsTipoGuia!TTipoEmision
    End If
    
    'Chequea Datos
    If txtClienteGuia.Caption = "" Then MsgBox "Ingrese el Cliente", vbExclamation, sMensaje: Exit Sub
    If txtTransporte.Caption = "" Then MsgBox "Ingrese el Transporte", vbExclamation, sMensaje: Exit Sub
    If nFItemGuia = 0 Then
         MsgBox "Configure la cantidad de Items de Impresion en Administracion - Parametros generales - Otros Datos - Configuracion Guias...", vbExclamation, sMensaje: Exit Sub
    End If
                   
    Set RsCantidad = Lib.OpenRecordset("select count(tTienda) as Codigo From " & sTemporal & " where ESTADO='G' Group by tTienda", Cn)
    If RsCantidad.RecordCount > 1 Then
       MsgBox "Error: No se puede generar una Guia con distintas Tiendas" & Chr(13) & "Verifique los pedidos escogidos", vbCritical, sMensaje
       Exit Sub
    End If
            
    'Genera y Actualiza los Numero de Documento
    sSerie = RsTipoGuia!tSerie
    sPrefijo = RsTipoGuia!prefijo
    sCorrela = Lib.Correlativo(RsTipoGuia!tUltimoNumero, 9)
    sTipoEmision = RsTipoGuia!TTipoEmision
    sDocumento = sPrefijo & sSerie & sCorrela
    sImp = RsTipoGuia!timpresora
    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoEmision & "' and tCaja ='" & sCaja & "'"
    
    'Crea el temporal para el detalle
    Screen.MousePointer = vbHourglass
    sGuiaTemporal = dbTemporal(sCaja, 5, "tCodigoProducto", "nVarChar(7)", _
                                         "nPrecioVenta", "Float", _
                                         "nCantidad", "Float", _
                                         "nVenta", "Float", _
                                         "tDocumento", "nVarChar(20)")
                                                                                                                                                                                                                                                               
    Cn.Execute "insert into " & sGuiaTemporal & " (tcodigoProducto, nPrecioVenta, nCantidad, nVenta) " & _
               "select tCodigoProducto, nPrecioVenta, nCantidad, nVenta from DPEDIDO " & _
               "where tCodigoPedido in (select tCodigoPedido from " & sTemporal & " where Estado='G') and DPEDIDO.tEstadoItem = 'N'"
               
               
    i = Int(Calcular("select count(tCodigoProducto) as Codigo from (select tCodigoProducto from " & sGuiaTemporal & " Group by tCodigoProducto) as AA", Cn) / nFItemGuia) + 1
    Dim aDocumento() As String
    ReDim aDocumento(i) As String
    'Inserta Detalle de Documento en un temporal
    Set RsGuiaGrupoProducto = Lib.OpenRecordset("Select tCodigoProducto from DPEDIDO where tCodigoPedido in (select tCodigoPedido from " & sTemporal & " where Estado='G' Group by tCodigoPedido) Group by tCodigoProducto", Cn)
    
    If RsGuiaGrupoProducto.RecordCount > 0 Then
       RsGuiaGrupoProducto.MoveFirst
       i = 1
       j = 1
       aDocumento(j) = sDocumento
       'Inserta Detalle Agrupado por Productos
       Do While Not RsGuiaGrupoProducto.EOF
          If i <= nFItemGuia Then
             'Genera Detalle
             Cn.Execute " insert into DGUIATRANSPORTE (tGuiaTransporte, tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tDocumento) " & _
                        " select '" & sDocumento & "', '" & Mid("000", 1, 3 - Len(LTrim(str(i)))) + LTrim(str(i)) & "', tCodigoProducto, avg(nPrecioVenta), sum(nCantidad), sum(nVenta), '' " & _
                        " from " & sGuiaTemporal & " where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "' " & _
                        " Group by tCodigoProducto"
                        
             'Graba el numero de Guia en el Temporal
             Cn.Execute "update " & sGuiaTemporal & " set tDocumento = '" & sDocumento & "' where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "'"
             
             'Graba el número de Guia en el Pedido
             Cn.Execute "update DPEDIDO set tGuiaTransporte = '" & sDocumento & "' where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "' and tCodigoPedido in (select tCodigoPedido from " & sTemporal & " where Estado='G') "
          Else
             'Genera Cabecera
             Cn.Execute " Insert into MGUIATRANSPORTE (tGuiaTransporte, fFecha, tCodigoDelivery, tDestinatario, tTransportista, tTienda, tUnidadTransporte, tDocumento, tEstadoGuia, nTotal, tUsuario, fRegistro, tCaja, tTurno, tMotivoTraslado,tTipoDocumento)" & _
                        " select '" & sDocumento & "', getdate(), '" & sCtaCte & "', '" & sCliente & "', '" & sTransportista & "', '" & sTienda & "', '" & sVehiculo & "', '', '01', sum(nVenta), '" & sUsuario & "', getdate(), '" & sCaja & "', '" & sTurno & "', '" & txtMotivoTraslado.Text & "', '" & sTipoDocumento & "'" & _
                        " from DGUIATRANSPORTE where tGuiaTransporte = '" & sDocumento & "' " & _
                        " Group by tGuiaTransporte "
          
             i = 1
             j = j + 1
             'Genera y Actualiza los Numero de Documento
             RsTipoGuia.Requery
             sSerie = RsTipoGuia!tSerie
             sCorrela = Lib.Correlativo(RsTipoGuia!tUltimoNumero, 9)
             sPrefijo = RsTipoGuia!prefijo
             sDocumento = sPrefijo & sSerie & sCorrela
             aDocumento(j) = sDocumento
             Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoEmision & "' and tCaja ='" & sCaja & "'"
             
             'Genera Detalle
             Cn.Execute " insert into DGUIATRANSPORTE (tGuiaTransporte, tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tDocumento) " & _
                        " select '" & sDocumento & "', '" & Mid("000", 1, 3 - Len(LTrim(str(i)))) + LTrim(str(i)) & "', tCodigoProducto, avg(nPrecioVenta), sum(nCantidad), sum(nVenta), '' " & _
                        " from " & sGuiaTemporal & " where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "' " & _
                        " Group by tCodigoProducto"
                        
             'Graba el numero de Guia en el Temporal
             Cn.Execute "update " & sGuiaTemporal & " set tDocumento = '" & sDocumento & "' where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "'"
                        
             'Graba el número de Guia en el Pedido
             Cn.Execute "update DPEDIDO set tGuiaTransporte = '" & sDocumento & "' where tCodigoProducto='" & RsGuiaGrupoProducto!tCodigoProducto & "' and tCodigoPedido in (select tCodigoPedido from " & sTemporal & " where Estado='G') "
          End If
          i = i + 1
          RsGuiaGrupoProducto.MoveNext
       Loop
         
       'Genera Cabecera
       Cn.Execute " Insert into MGUIATRANSPORTE (tGuiaTransporte, fFecha, tCodigoDelivery, tDestinatario, tTransportista, tTienda, tUnidadTransporte, tDocumento, tEstadoGuia, nTotal, tUsuario, fRegistro, tCaja, tTurno, tMotivoTraslado, tTipoDocumento)" & _
                  " select '" & sDocumento & "', getdate(), '" & sCtaCte & "', '" & sCliente & "', '" & sTransportista & "', '" & sTienda & "', '" & sVehiculo & "', '', '01', sum(nVenta), '" & sUsuario & "', getdate(), '" & sCaja & "', '" & sTurno & "', '" & txtMotivoTraslado.Text & "' , '" & sTipoDocumento & "'" & _
                  " from DGUIATRANSPORTE where tGuiaTransporte = '" & sDocumento & "' " & _
                  " Group by tGuiaTransporte "
                  
    End If
        
    'Imprime Guias
    For i = 1 To j
        If lAlmacen And sEmpresa = "100" Then
            'CASO CORP.FERNANDEZ
            If BDLink = "1" Then
                Isql = "select vg.*, isnull(at.UnidadEntrada,'Sin UM') as UM from vGuiaTransporte vg left join " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.vPRODUCTO at on vg.tenlace= at.tCodigoProducto where  vg.tGuiaTransporte ='" & aDocumento(i) & "'"
            Else
                Isql = "select vg.*, isnull(at.UnidadEntrada,'Sin UM') as UM from vGuiaTransporte vg left join " & sAlmacenMDB & ".dbo.vPRODUCTO at on vg.tenlace= at.tCodigoProducto where  vg.tGuiaTransporte ='" & aDocumento(i) & "'"
            End If
        Else
             '24092018 CLIENTE : DUZKA
            If sEmpresa = "012" Then
                Isql = "SELECT vGuiaTransporte.*,PedidoObservacion,'Sin UM' as UM, fEntrega FROM [vGuiaTransporte]  "
                Isql = Isql & " left join (SELECT top 1 tguiatransporte,ISNULL(MPEDIDO.tObservacion,'') as PedidoObservacion, isnull(MPEDIDO.fProgramacion,getdate()) as fEntrega FROM MPEDIDO inner join dpedido on MPEDIDO.tCodigoPedido=dpedido.tCodigoPedido "
                Isql = Isql & " WHERE dpedido.tguiatransporte='" & aDocumento(i) & "') PEDIDO "
                Isql = Isql & " on [vGuiaTransporte].tGuiaTransporte =PEDIDO.tGuiaTransporte "
                Isql = Isql & " where vGuiaTransporte.tGuiaTransporte ='" & aDocumento(i) & "'"
                
            Else
                Isql = "SELECT *,'Sin UM' as UM from vGuiaTransporte where vGuiaTransporte.tGuiaTransporte ='" & aDocumento(i) & "' "
            End If
        End If
'
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
    
        If RsImpresion.RecordCount = 0 Then
           LimpiaRs
           MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
        Else
           'Configura la Impresora
           Imprimir (sImp)
           Printer.FontName = sFont
           Printer.FontBold = False

            If lImprimirGuiaFormatoCorp Then
                Genera4 "N"
                Dim path As String
                path = ""
                
                RsNewRptGuia.xgaaa = "C:\Documents and Settings\Erick\Mis documentos\Mis imágenes\logo fernandez.jpg"
                RsNewRptGuia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                RsNewRptGuia.PaperOrientation = crPortrait
                
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
            Else
                ImprimeGuiaTransporte RsImpresion, sEmpresa
            End If
           
           
           On Error GoTo fin
            Ruta = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
            If Dir(Ruta, vbDirectory) <> "" Then
            clasexml.CrearXMLGuia RsImpresion, Ruta, aDocumento(i), ""
           End If
fin:
           LimpiaRs

               'CESAR FACTURACION ELECTRONICA
               If pais = "002" Then
                   If lFacturacionE And lFEEcuador = False Then
                       If IIf(RsTipoGuia!lFacturacionElectronica = True, 1, 0) Then
                            'CABECERA
                            Dim oComando As clsComando
                            Set oComando = New clsComando
                            If Not oComando.CreateCmdSp("USP_InsertaFactGuia", Cn) Then
                                 Set oComando = Nothing
                                 Exit Sub
                            End If
                            oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, aDocumento(i)
            
                            If Not oComando.GetParamOK Then
                                 Set oComando = Nothing
                                 Exit Sub
                            End If
                            If Not oComando.ExecSP Then
                                 Set oComando = Nothing
                                 Exit Sub
                            End If
                            'DETALLE
                            Dim oComandoDetalle As clsComando
                            Set oComandoDetalle = New clsComando
                            If Not oComandoDetalle.CreateCmdSp("USP_InsertaFactGuiaDetalle", Cn) Then
                                 Set oComandoDetalle = Nothing
                                 Exit Sub
                            End If
                            oComandoDetalle.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, aDocumento(i)
            
                            If Not oComandoDetalle.GetParamOK Then
                                 Set oComandoDetalle = Nothing
                                 Exit Sub
                            End If
                            If Not oComandoDetalle.ExecSP Then
                                 Set oComandoDetalle = Nothing
                                 Exit Sub
                            End If
                       End If
                   End If
               End If
               '---------------------------------
           
           
           
           
           
       End If
   Next i
          
   'Borra lo facturado
   Cn.Execute "delete from " & sTemporal
   Isql = "Insert into " & sTemporal & _
          " SELECT ' ' as Estado, MPEDIDO.tCodigoPedido, MPEDIDO.fFecha, Sum(DPEDIDO.nPrecioNeto) AS nNeto, Sum(DPEDIDO.nImpuesto1) AS nImpuesto1, Sum(DPEDIDO.nImpuesto2) AS nImpuesto2, Sum(DPEDIDO.nImpuesto3) AS nImpuesto3, Sum(DPEDIDO.nVenta) AS nVenta, Count(DPEDIDO.tItem) AS nItem, tClienteDelivery, tTienda, tGuiaTransporte " & _
          "FROM MPEDIDO LEFT JOIN DPEDIDO ON MPEDIDO.tCodigoPedido = DPEDIDO.tCodigoPedido " & _
          "where MPEDIDO.tClienteCtaCte = '" & sCtaCte & "' and DPEDIDO.tEstadoItem = 'N' and MPEDIDO.tEstadoPedido = '04' and isnull(DPEDIDO.tGuiaTransporte,'')='' " & _
          "GROUP BY MPEDIDO.fFecha, MPEDIDO.tClienteCtaCte, MPEDIDO.tCodigoPedido, MPEDIDO.tClienteDelivery, tTienda, tGuiaTransporte " & _
          "Union " & _
          "SELECT '' as Estado, '' as tCodigoPedido, dbo.MGUIATRANSPORTE.fFecha, SUM(dbo.DGUIATRANSPORTE.nPrecioVenta) AS nPrecioVenta, 0 AS nImpuesto1, 0 AS nImpuesto2, 0 AS nImpuesto3, SUM(dbo.DGUIATRANSPORTE.nVenta) AS nVenta, COUNT(dbo.DGUIATRANSPORTE.tItem) AS tItem, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda , dbo.MGUIATRANSPORTE.tGuiaTransporte " & _
          "FROM dbo.MGUIATRANSPORTE INNER JOIN dbo.DGUIATRANSPORTE ON dbo.MGUIATRANSPORTE.tGuiaTransporte = dbo.DGUIATRANSPORTE.tGuiaTransporte " & _
          "where MGUIATRANSPORTE.tCodigoDelivery = '" & sCtaCte & "' and MGUIATRANSPORTE.tEstadoGuia = '01' " & _
          "GROUP BY dbo.MGUIATRANSPORTE.fFecha, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda, dbo.MGUIATRANSPORTE.tGuiaTransporte "
                    
   Cn.Execute Isql
   Cn.Execute "update " & sTemporal & " set Estado='F' where isnull(tGuiaTransporte,'') <> '' "
   
   RsDetalle1.Requery
   RsDetalle2.Requery
   RsDetalle3.Requery
   RsDetalle4.Requery
   RsDetalle5.Requery
   RsDetalle6.Requery
   CalculaMonto
        
   'Borra el temporal
   Cn.Execute "drop table " & sGuiaTemporal
   
   Screen.MousePointer = vbDefault

End Sub
Public Sub Genera4(ByVal Opcion As String)
'--------CESAR-------Registro Sunat
    Dim oComando As clsComando
           
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
   '--------Emite el reporte guia
     
    If Not oComando.CreateCmdSp("usp_Inforest_ObtieneRptGuia", Cn) Then
      Set oComando = Nothing
      Exit Sub
    End If
       
    If Opcion = "R" Then
        oComando.CreateParameter "@tGuiaTransporte", adVarChar, adParamInput, 50, RsDetalle5!tGuiaTransporte ' sDocumento
    Else
        oComando.CreateParameter "@tGuiaTransporte", adVarChar, adParamInput, 50, sDocumento
    End If
   
    If BDLink = "1" Then
        oComando.CreateParameter "@Almacen", adVarChar, adParamInput, 100, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
    Else
        oComando.CreateParameter "@Almacen", adVarChar, adParamInput, 100, sAlmacenMDB  'CASO CORP.FERNANDEZ
    End If
    
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
     RsNewRptGuia.DiscardSavedData
     RsNewRptGuia.Database.SetDataSource RsPrinter
        
     frmEmite.CRViewer.ReportSource = RsNewRptGuia
    
     Screen.MousePointer = vbDefault
    
    '/////////////////////////////////
End Sub
Private Sub Form_Load()
    Centrar Me
   
   'Impresoras
'   If pais = "002" Then 'Ecuador
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' order by tTipoEmision", Cn)
'   Else
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 order by tTipoEmision", Cn)
'   End If

    SSTab1.TabCaption(1) = "Generar " & Guia
    SSTab1.TabCaption(2) = Guia & " a Facturar"

    grdDetalle3.Caption = "Pendientes por Generar " & Guia
    grdDetalle4.Caption = "Pedidos a Generar " & Guia
    grdDetalle5.Caption = Guia & " por Facturar"
    grdDetalle6.Caption = Guia & " a Facturar"
    
    cmdOpcion(8).Caption = "Anular " & Guia
    cmdOpcion(10).Caption = "Imp. " & Guia
    
    Label(1).Caption = "Total de Pendientes por Generar " & Guia & " :"
    Label(4).Caption = "Total a Generar " & Guia & " :"
    
    Frame6.Caption = " Tipos de " & Guia

    If pais = "000" Then
        If lFacturacionE Then
            If lFESpring Then
                SSTab1.TabEnabled(2) = False
            End If
        End If
    End If

   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If
   
      
   If RsTipoDocumento.RecordCount = 0 Then
      MsgBox "No se ha ingresado los Documentos por Caja", vbCritical, "Error Fatal"
      Unload Me
   End If
   sCtaCte = frmCtaCte.RsCabecera!codigo
   sCliente = ""
   txtCtaCte.Caption = frmCtaCte.grdGrilla.Columns(1).Text
      
   Isql = "select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=1 And lActivo = 1 And lFacturacionElectronica=0 UNION select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=1 And lActivo = 1 And lFacturacionElectronica=1"
   Set RsTipoGuia = Lib.OpenRecordset(Isql, Cn)
   
   'Detalle 1
   sTemporal = dbTemporal(sCaja, 12, "Estado", "nVarChar(1)", _
                                     "tCodigoPedido", "nVarChar(10)", _
                                     "fFecha", "smalldatetime", _
                                     "nNeto", "Float", _
                                     "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "nItem", "Float", _
                                     "tCodigoDelivery", "nVarChar(7)", _
                                     "tTienda", "nVarChar(3)", _
                                     "tGuiaTransporte", "nVarChar(20)")
  
   Isql = "Insert into " & sTemporal & _
          " SELECT ' ' as Estado, MPEDIDO.tCodigoPedido, MPEDIDO.fFecha, Sum(DPEDIDO.nPrecioNeto) AS nNeto, Sum(DPEDIDO.nImpuesto1) AS nImpuesto1, Sum(DPEDIDO.nImpuesto2) AS nImpuesto2, Sum(DPEDIDO.nImpuesto3) AS nImpuesto3, Sum(DPEDIDO.nVenta)+sum(case when vproducto.lley1=1 then dpedido.ncantidad*(" & CDbl(sValorLey1) & ") else 0 end  ) AS nVenta, Count(DPEDIDO.tItem) AS nItem, tClienteDelivery, tTienda, tGuiaTransporte " & _
          "FROM MPEDIDO LEFT JOIN DPEDIDO ON MPEDIDO.tCodigoPedido = DPEDIDO.tCodigoPedido left join vproducto on dpedido.tcodigoproducto=vproducto.codigo " & _
          "where MPEDIDO.tClienteCtaCte = '" & sCtaCte & "' and DPEDIDO.tEstadoItem = 'N' and MPEDIDO.tEstadoPedido = '04' and isnull(DPEDIDO.tGuiaTransporte,'')='' " & _
          "GROUP BY MPEDIDO.fFecha, MPEDIDO.tClienteCtaCte, MPEDIDO.tCodigoPedido, MPEDIDO.tClienteDelivery, tTienda, tGuiaTransporte " & _
          "Union " & _
          "SELECT '' as Estado, '' as tCodigoPedido, dbo.MGUIATRANSPORTE.fFecha, SUM(dbo.DGUIATRANSPORTE.nPrecioVenta) AS nPrecioVenta, 0 AS nImpuesto1, 0 AS nImpuesto2, 0 AS nImpuesto3, SUM(dbo.DGUIATRANSPORTE.nVenta) AS nVenta, COUNT(dbo.DGUIATRANSPORTE.tItem) AS tItem, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda , dbo.MGUIATRANSPORTE.tGuiaTransporte " & _
          "FROM dbo.MGUIATRANSPORTE INNER JOIN dbo.DGUIATRANSPORTE ON dbo.MGUIATRANSPORTE.tGuiaTransporte = dbo.DGUIATRANSPORTE.tGuiaTransporte " & _
          "where MGUIATRANSPORTE.tCodigoDelivery = '" & sCtaCte & "' and MGUIATRANSPORTE.tEstadoGuia = '01' " & _
          "GROUP BY dbo.MGUIATRANSPORTE.fFecha, dbo.MGUIATRANSPORTE.tCodigoDelivery, dbo.MGUIATRANSPORTE.tTienda, dbo.MGUIATRANSPORTE.tGuiaTransporte "
                    
   Cn.Execute Isql
   Cn.Execute "update " & sTemporal & " set Estado='F' where isnull(tGuiaTransporte,'') <> '' "
   
   'Primera Grilla
   Isql = "SELECT tCodigoPedido, fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = ' ' order by tCodigoPedido"
   Set RsDetalle1 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(5, grdDetalle1, "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Pedido", 2, "tCodigoPedido", 950, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00", _
                                   "Itm", 2, "nItem", 530, 1, 0, "#,##0")
                                   
   'Segunda Grilla
   Isql = "SELECT tCodigoPedido, fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = 'X' order by tCodigoPedido"
   Set RsDetalle2 = Lib.OpenRecordset(Isql, Cn)

   Call ConfGrilla(5, grdDetalle2, "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Pedido", 2, "tCodigoPedido", 950, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00", _
                                   "Itm", 2, "nItem", 530, 1, 0, "#,##0")
                    
   'Tercera Grilla
   Isql = "SELECT tCodigoPedido, fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = ' ' order by tCodigoPedido"
   Set RsDetalle3 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(5, grdDetalle3, "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Pedido", 2, "tCodigoPedido", 950, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00", _
                                   "Items", 2, "nItem", 530, 1, 0, "#,##0")
   
   'Cuarta Grilla
   Isql = "SELECT tCodigoPedido, fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = 'G' order by tCodigoPedido"
   Set RsDetalle4 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(5, grdDetalle4, "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Pedido", 2, "tCodigoPedido", 950, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00", _
                                   "Items", 2, "nItem", 530, 1, 0, "#,##0")
   
   'Quinta Grilla
   Isql = "SELECT fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion, tGuiaTransporte " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = 'F' order by tGuiaTransporte"
   Set RsDetalle5 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(4, grdDetalle5, "Guia", 2, "tGuiaTransporte", 1500, 0, 0, "", _
                                   "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00")
   
   'Sexta Grilla
   Isql = "SELECT fFecha, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nItem, Descripcion, tGuiaTransporte " & _
          "FROM [" & sTemporal & "] LEFT OUTER JOIN dbo.vTienda ON [" & sTemporal & "].tCodigoDelivery = dbo.vTienda.tCodigoDelivery and [" & sTemporal & "].tTienda=dbo.vTienda.Codigo " & _
          "where Estado = 'A' order by tGuiaTransporte"
   Set RsDetalle6 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(4, grdDetalle6, "Guia", 2, "tGuiaTransporte", 1550, 0, 0, "", _
                                   "Tienda", 2, "Descripcion", 1350, 0, 0, "", _
                                   "Fecha", 2, "fFecha", 800, 0, 0, "dd/mm/yy", _
                                   "Monto", 2, "nVenta", 800, 1, 0, "#,###,##0.00")
                                         
   Set grdDetalle1.DataSource = RsDetalle1
   Set grdDetalle2.DataSource = RsDetalle2
   Set grdDetalle3.DataSource = RsDetalle3
   Set grdDetalle4.DataSource = RsDetalle4
   Set grdDetalle5.DataSource = RsDetalle5
   Set grdDetalle6.DataSource = RsDetalle6
   
   
'   txtLinea.Text = Format(frmCtaCte.grdGrilla.Columns(3).Text, "###,##0.00")
   
   Set RsCliente = Lib.OpenRecordset("select * from vCliente", Cn)
   Set RsCtaCte = Lib.OpenRecordset("select * from vCompania", Cn)
   RsCtaCte.MoveFirst
   RsCtaCte.Find ("Codigo='" & sCtaCte & "'")
   txtCliente.Caption = ""
   txtClienteGuia.Caption = ""
   
   If IsNull(RsCtaCte!tCodigoCliente) Or RsCtaCte!tCodigoCliente <> "" Then
        If RsCliente.RecordCount > 0 Then
           RsCliente.MoveFirst
           RsCliente.Find ("Codigo='" & RsCtaCte!tCodigoCliente & "'")
           If Not (RsCliente.EOF) Then
                sTemp = IIf(IsNull(RsCliente!tIdentidad), "", RsCliente!tIdentidad)
                txtCliente.Caption = IIf(IsNull(RsCliente!descripcion), "", RsCliente!descripcion)
                txtClienteGuia.Caption = IIf(IsNull(RsCliente!descripcion), "", RsCliente!descripcion)
                txtClienteGuiaFactura.Caption = IIf(IsNull(RsCliente!descripcion), "", RsCliente!descripcion)
                sCliente = RsCtaCte.Fields("tcodigocliente")
           End If
        End If
   End If
   
   AsignaComando 4, RsTipoDocumento, cmdTipoDocumento()
   AsignaComando 4, RsTipoGuia, cmdTipoGuia()
   AsignaComando 4, RsTipoDocumento, cmdTipoDocumento2()
  
  'Imprime el Totales
   CalculaMonto
   SSTab1.Tab = 0
   Screen.MousePointer = vbDefault

End Sub

Public Sub CalculaMonto()
   txtTotal1.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = ' '", Cn), "###,###,###,##0.00")
   txtTotal2.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'X'", Cn), "###,###,###,##0.00")
   txtTotal3.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = ' '", Cn), "###,###,###,##0.00")
   txtTotal4.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'G'", Cn), "###,###,###,##0.00")
   Me.Label13.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'F'", Cn), "###,###,###,##0.00")
   Me.Label11.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'A'", Cn), "###,###,###,##0.00")
End Sub

Private Sub cmdMovimiento_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
           Case Is = 0  ' Pasar Todas
                Cn.Execute "update " & sTemporal & " set Estado ='X' where Estado= ' '"
           
           Case Is = 1  ' Pasar una
                If RsDetalle1.RecordCount <> 0 Then
                  Cn.Execute "update " & sTemporal & " set Estado ='X' where tCodigoPedido = '" & RsDetalle1!tCodigoPedido & "'"
                End If
           
           Case Is = 2  ' Quitar una
                If RsDetalle2.RecordCount <> 0 Then
                   Cn.Execute "update " & sTemporal & " set Estado =' ' where tCodigoPedido = '" & RsDetalle2!tCodigoPedido & "'"
                End If
                
           Case Is = 3 ' Quitar Todas
                Cn.Execute "update " & sTemporal & " set Estado =' ' where Estado='X'"
                
                
           Case Is = 4  ' Pasar Todas
                Cn.Execute "update " & sTemporal & " set Estado ='G' where Estado= ' '"
           
           Case Is = 5  ' Pasar una
                If RsDetalle1.RecordCount <> 0 Then
                  Cn.Execute "update " & sTemporal & " set Estado ='G' where tCodigoPedido = '" & RsDetalle3!tCodigoPedido & "'"
                End If
           
           Case Is = 6  ' Quitar una
                If RsDetalle4.RecordCount <> 0 Then
                   Cn.Execute "update " & sTemporal & " set Estado =' ' where tCodigoPedido = '" & RsDetalle4!tCodigoPedido & "'"
                End If
                
           Case Is = 7 ' Quitar Todas
                Cn.Execute "update " & sTemporal & " set Estado =' ' where Estado = 'G'"
                
                
           'Guias x Factura
           Case Is = 8  ' Pasar Todas
                Cn.Execute "update " & sTemporal & " set Estado ='A' where Estado= 'F'"
           
           Case Is = 9  ' Pasar una
                If RsDetalle5.RecordCount <> 0 Then
                  Cn.Execute "update " & sTemporal & " set Estado ='A' where tGuiaTransporte = '" & RsDetalle5!tGuiaTransporte & "'"
                End If
           
           Case Is = 10  ' Quitar una
                If RsDetalle6.RecordCount <> 0 Then
                   Cn.Execute "update " & sTemporal & " set Estado ='F' where tGuiaTransporte = '" & RsDetalle6!tGuiaTransporte & "'"
                End If
                
           Case Is = 11 ' Quitar Todas
                Cn.Execute "update " & sTemporal & " set Estado ='F' where Estado = 'A'"
                
    End Select
    
    RsDetalle1.Requery
    RsDetalle2.Requery
    RsDetalle3.Requery
    RsDetalle4.Requery
    RsDetalle5.Requery
    RsDetalle6.Requery
    CalculaMonto
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0  ' Primero Detalle1
                MoverPuntero Primero, grdDetalle1
           Case Is = 1  ' PgUp Detalle1
                MoverPuntero pgup, grdDetalle1
           Case Is = 2  ' Previo Detalle1
                MoverPuntero previo, grdDetalle1
           Case Is = 3  ' Siguiente Detalle1
                MoverPuntero siguiente, grdDetalle1
           Case Is = 4  ' Ultimo Detalle1
                MoverPuntero pgdn, grdDetalle1
           Case Is = 5  ' Ultimo Detalle1
                MoverPuntero Ultimo, grdDetalle1
                
           Case Is = 6  ' Primero Detalle2
                MoverPuntero Primero, grdDetalle2
           Case Is = 7  ' PgUp Detalle2
                MoverPuntero pgup, grdDetalle2
           Case Is = 8  ' Previo Detalle2
                MoverPuntero previo, grdDetalle2
           Case Is = 9  ' Siguiente Detalle2
                MoverPuntero siguiente, grdDetalle2
           Case Is = 10 ' Ultimo Detalle2
                MoverPuntero pgdn, grdDetalle2
           Case Is = 11 ' Ultimo Detalle2
                MoverPuntero Ultimo, grdDetalle2
                
           Case Is = 12  ' Primero Detalle3
                MoverPuntero Primero, grdDetalle3
           Case Is = 13  ' PgUp Detalle3
                MoverPuntero pgup, grdDetalle3
           Case Is = 14  ' Previo Detalle3
                MoverPuntero previo, grdDetalle3
           Case Is = 15  ' Siguiente Detalle3
                MoverPuntero siguiente, grdDetalle3
           Case Is = 16  ' Ultimo Detalle3
                MoverPuntero pgdn, grdDetalle3
           Case Is = 17  ' Ultimo Detalle3
                MoverPuntero Ultimo, grdDetalle3
           
           Case Is = 18  ' Primero Detalle4
                MoverPuntero Primero, grdDetalle4
           Case Is = 19  ' PgUp Detalle4
                MoverPuntero pgup, grdDetalle4
           Case Is = 20  ' Previo Detalle4
                MoverPuntero previo, grdDetalle4
           Case Is = 21  ' Siguiente Detalle4
                MoverPuntero siguiente, grdDetalle4
           Case Is = 22  ' Ultimo Detalle4
                MoverPuntero pgdn, grdDetalle4
           Case Is = 23  ' Ultimo Detalle4
                MoverPuntero Ultimo, grdDetalle4
                                
           Case Is = 24  ' Primero Detalle5
                MoverPuntero Primero, grdDetalle5
           Case Is = 25  ' PgUp Detalle5
                MoverPuntero pgup, grdDetalle5
           Case Is = 26  ' Previo Detalle5
                MoverPuntero previo, grdDetalle5
           Case Is = 27  ' Siguiente Detalle5
                MoverPuntero siguiente, grdDetalle5
           Case Is = 28  ' Ultimo Detalle5
                MoverPuntero pgdn, grdDetalle5
           Case Is = 29  ' Ultimo Detalle5
                MoverPuntero Ultimo, grdDetalle5
           
           Case Is = 30  ' Primero Detalle6
                MoverPuntero Primero, grdDetalle6
           Case Is = 31  ' PgUp Detalle6
                MoverPuntero pgup, grdDetalle6
           Case Is = 32  ' Previo Detalle6
                MoverPuntero previo, grdDetalle6
           Case Is = 33  ' Siguiente Detalle6
                MoverPuntero siguiente, grdDetalle6
           Case Is = 34  ' Ultimo Detalle6
                MoverPuntero pgdn, grdDetalle6
           Case Is = 35  ' Ultimo Detalle6
                MoverPuntero Ultimo, grdDetalle6
                
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cn.Execute "drop table " & sTemporal

    Set RsTipoDocumento = Nothing
    Set RsDetalle1 = Nothing
    Set RsDetalle2 = Nothing
    Set RsCtaCte = Nothing
    Set RsImpresion = Nothing
    Set RsRapido = Nothing
    Set frmCtaCteDetalle = Nothing
End Sub

'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
                fImpresionDiaContable = obtieneDiaContable
End Function

Private Sub RevierteDocumento(ByVal RsDatos As Recordset)
    Dim UltimoCorrelativo As String
    Dim TipoDocumento As String
    Dim DocEliminar As String
    
    RsDatos.Requery
    RsDatos.MoveFirst
    Do While Not RsDatos.EOF
        DocEliminar = Calcular("select top 1 (D.tDocumento) as Codigo from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & RsDatos!tCodigoPedido & "' and M.tEstadoDocumento <> '04'", Cn)
        TipoDocumento = Calcular("select tTipoDocumento as Codigo from mDocumento where tDocumento = '" & DocEliminar & "'", Cn)
        Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & RsDatos!tCodigoPedido & "'"
        Cn.Execute "update MPEDIDO set tEstadoPedido = '04' where tCodigoPedido = '" & RsDatos!tCodigoPedido & "'"
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & RsDatos!tCodigoPedido & "' and  tDocumento ='" & DocEliminar & "'"
        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        
        UltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & TipoDocumento & "'", Cn)
        UltimoCorrelativo = Right(UltimoCorrelativo, 9)
        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & UltimoCorrelativo & "' where tTipoEmision ='" & TipoDocumento & "' and tCaja ='" & sCaja & "'"
     
        RsDetalle2.MoveNext
    Loop
                                                       
End Sub


Private Sub RevierteDocumentoGUIA(ByVal RsDatos As Recordset)
    Dim UltimoCorrelativo As String
    Dim TipoDocumento As String
    Dim DocEliminar As String
    
    RsDatos.Requery
    RsDatos.MoveFirst
    Do While Not RsDatos.EOF
        DocEliminar = Calcular("select top 1 (Documento) as Codigo from vGuiaTransporte where tGuiaTransporte = '" & RsDatos!tGuiaTransporte & "'", Cn)
        TipoDocumento = Calcular("select tTipoDocumento as Codigo from mDocumento where tDocumento = '" & DocEliminar & "'", Cn)

        Cn.Execute "Update MGUIATRANSPORTE Set tEstadoGuia='01', tDocumento='' Where tGuiaTransporte='" & RsDatos!tGuiaTransporte & "'"
        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & DocEliminar & "'"
        Cn.Execute "update dpedido set tdocumento='',tfacturado='' Where tguiatransporte= '" & RsDatos!tGuiaTransporte & "'"
        
        UltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & TipoDocumento & "'", Cn)
        UltimoCorrelativo = Right(UltimoCorrelativo, 9)
        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & UltimoCorrelativo & "' where tTipoEmision ='" & TipoDocumento & "' and tCaja ='" & sCaja & "'"
     
        RsDatos.MoveNext
    Loop
                                                       
End Sub





