VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNuevoDelivery 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Cliente Frecuente"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmNuevoDelivery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Tiendas"
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
      Index           =   9
      Left            =   180
      Picture         =   "frmNuevoDelivery.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7740
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "frmNuevoDelivery.frx":0404
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtApellido"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNombre"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDireccion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTelefono"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtReferencia"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtObservacion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdOpcion(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdOpcion(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdOpcion(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdOpcion(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdOpcion(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdOpcion(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraZona"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame7"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmNuevoDelivery.frx":0420
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dtpfecha"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtemail"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdOpcion(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtPuntos"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdOpcion(11)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdOpcion(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdOpcion(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtDescuento"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Cliente / Cuentas Corrientes"
      TabPicture(2)   =   "frmNuevoDelivery.frx":043C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtLinea"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkCtaCte"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdLinea"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "BtnBuscarCentral"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.CommandButton BtnBuscarCentral 
         Caption         =   "Buscar Central"
         Height          =   555
         Left            =   -66240
         TabIndex        =   74
         Top             =   2040
         Width           =   1475
      End
      Begin VB.Frame Frame7 
         Caption         =   " Tipo de Cliente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   180
         TabIndex        =   70
         Top             =   6300
         Width           =   10185
         Begin VB.CommandButton cmdBuscaTipoCliente 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":0458
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtTipoClienteDetalle 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   397
            Width           =   6795
         End
         Begin VB.TextBox txtTipoCliente 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   300
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   397
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmdLinea 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -70200
         Picture         =   "frmNuevoDelivery.frx":089A
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   960
         Width           =   1275
      End
      Begin VB.Frame Frame6 
         Caption         =   " Estado "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   -74760
         TabIndex        =   58
         Top             =   5280
         Width           =   10185
         Begin VB.CommandButton cmdEstado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":0CDC
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   383
            Width           =   1275
         End
         Begin VB.TextBox txtEstadoCodigo 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txtEstadoDescripcion 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1980
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   420
            Width           =   6675
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Identidad "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   -74760
         TabIndex        =   55
         Top             =   1560
         Width           =   10185
         Begin VB.CommandButton cmdTipoIdentidad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2880
            Picture         =   "frmNuevoDelivery.frx":111E
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   443
            Width           =   1275
         End
         Begin VB.TextBox txtTipoIdentidad 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   480
            Width           =   2595
         End
         Begin VB.CommandButton cmdNumeroIdentidad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7200
            Picture         =   "frmNuevoDelivery.frx":1560
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   480
            Width           =   1275
         End
         Begin VB.TextBox txtNumeroIdentidad 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Tipo de Sub Cuenta Corriente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   -74760
         TabIndex        =   50
         Top             =   4060
         Width           =   10185
         Begin VB.TextBox txtTipoSubCuentaCorrienteDescripcion 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1980
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   420
            Width           =   6675
         End
         Begin VB.TextBox txtTipoSubCuentaCorrienteCodigo 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   420
            Width           =   1515
         End
         Begin VB.CommandButton cmdTipoSubCuentaCorriente 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":19A2
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   383
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Tipo Cuenta Corriente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   -74760
         TabIndex        =   46
         Top             =   2840
         Width           =   10185
         Begin VB.TextBox txtTipoCuentaCorrienteDetalle 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   420
            Width           =   6735
         End
         Begin VB.TextBox txtTipoCuentaCorrienteCodigo 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   420
            Width           =   1515
         End
         Begin VB.CommandButton cmdTipoCuentaCorriente 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":1DE4
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   383
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkCtaCte 
         Caption         =   "Habilitar Cuenta Corriente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74820
         TabIndex        =   45
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtLinea 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "1,500.00"
         Top             =   1020
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tarjeta de Crédito "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   -74820
         TabIndex        =   35
         Top             =   3360
         Width           =   10185
         Begin VB.TextBox txtNumero 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4680
            TabIndex        =   39
            Top             =   480
            Width           =   3720
         End
         Begin VB.CommandButton cmdBuscarTarjeta 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3120
            Picture         =   "frmNuevoDelivery.frx":2226
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   443
            Width           =   1275
         End
         Begin VB.TextBox txtTarjeta 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   300
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   480
            Width           =   2655
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Numero"
            Height          =   555
            Index           =   9
            Left            =   8580
            TabIndex        =   36
            Top             =   443
            Width           =   1275
         End
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   -71700
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   660
         Width           =   2535
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Des&cuento"
         Height          =   555
         Index           =   8
         Left            =   -68640
         TabIndex        =   25
         Top             =   623
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Puntos"
         Height          =   555
         Index           =   10
         Left            =   -68640
         TabIndex        =   23
         Top             =   1263
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Fecha"
         Height          =   555
         Index           =   11
         Left            =   -68640
         TabIndex        =   22
         Top             =   1903
         Width           =   1275
      End
      Begin VB.TextBox txtPuntos 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   -71700
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Si"
         Top             =   1300
         Width           =   2535
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Email"
         Height          =   555
         Index           =   12
         Left            =   -66240
         TabIndex        =   20
         Top             =   2543
         Width           =   1275
      End
      Begin VB.TextBox txtemail 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -71700
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2580
         Width           =   5235
      End
      Begin VB.Frame fraZona 
         Caption         =   " Zona "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   240
         TabIndex        =   18
         Top             =   3900
         Width           =   10185
         Begin VB.CommandButton cmdBuscarZona 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":2668
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   383
            Width           =   1275
         End
         Begin VB.TextBox txtZonaCodigo 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   300
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txtZonaDetalle 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   420
            Width           =   6795
         End
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Apellidos"
         Height          =   555
         Index           =   0
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Nombres"
         Height          =   555
         Index           =   1
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Dirección"
         Height          =   555
         Index           =   2
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Id."
         Height          =   555
         Index           =   3
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Observación"
         Height          =   555
         Index           =   7
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3270
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "&Referencia"
         Height          =   555
         Index           =   4
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2715
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Caption         =   " Distrito "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   240
         TabIndex        =   11
         Top             =   5100
         Width           =   10185
         Begin VB.TextBox txtDistritoCodigo 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   300
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   397
            Width           =   1515
         End
         Begin VB.TextBox txtDistritoDetalle 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   397
            Width           =   6795
         End
         Begin VB.CommandButton cmdBuscarDistrito 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8820
            Picture         =   "frmNuevoDelivery.frx":2AAA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3285
         Width           =   6750
      End
      Begin VB.TextBox txtReferencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2730
         Width           =   6750
      End
      Begin VB.TextBox txtTelefono 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   517
         Width           =   2535
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
         Height          =   525
         Left            =   2160
         MaxLength       =   80
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2175
         Width           =   6750
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1627
         Width           =   6750
      End
      Begin VB.TextBox txtApellido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1072
         Width           =   6750
      End
      Begin MSComCtl2.DTPicker dtpfecha 
         Height          =   480
         Left            =   -71700
         TabIndex        =   24
         Top             =   1940
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   847
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM"
         Format          =   54591491
         CurrentDate     =   38637
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Correo Electrónico :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74820
         TabIndex        =   66
         Top             =   2640
         Width           =   2955
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Nacimiento :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74820
         TabIndex        =   65
         Top             =   1963
         Width           =   2955
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Canje por Puntos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74280
         TabIndex        =   64
         Top             =   1323
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Aplica Descuento :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74340
         TabIndex        =   63
         Top             =   683
         Width           =   2475
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Línea :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74460
         TabIndex        =   54
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label10 
         Caption         =   "(Teléfono, Código de Socio o Código Interno Cliente)"
         Height          =   315
         Left            =   4800
         TabIndex        =   34
         Top             =   660
         Width           =   3735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Observación :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         TabIndex        =   33
         Top             =   3330
         Width           =   1755
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   32
         Top             =   2775
         Width           =   1515
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   31
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombres :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   30
         Top             =   1650
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellidos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   29
         Top             =   1095
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Id. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1500
         TabIndex        =   28
         Top             =   570
         Width           =   555
      End
      Begin VB.Label Label 
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
         Left            =   -71685
         TabIndex        =   27
         Top             =   2145
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
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
      Index           =   6
      Left            =   7980
      Picture         =   "frmNuevoDelivery.frx":2EEC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7740
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
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
      Index           =   5
      Left            =   9360
      Picture         =   "frmNuevoDelivery.frx":2FEE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7740
      Width           =   1275
   End
End
Attribute VB_Name = "frmNuevoDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCliente As String
Dim xsCliente As String

Dim sTarjeta As String
Dim nCodigo As String
Dim RsZona As Recordset
Dim RsDistrito As Recordset
Dim RsTipoCuentaCorriente As Recordset
Dim RsTipoSubCuentaCorriente As Recordset
Dim RsEstado As Recordset
Dim RsTipoIdentidad As Recordset
Dim RsTipoCliente As Recordset
Dim sTipoCuentaCorriente As String
Dim sTipoSubCuentaCorriente As String
Dim sEstado As String
Dim sTipoIdentidad As String
Dim sZona As String
Dim sDistrito As String
Dim sTipoCliente As String
Dim i As Integer
Dim RsDelivery As Recordset

Private Sub BtnBuscarCentral_Click()
If Trim(txtNumeroIdentidad.Text) = "" Then
    MsgBox "Debe Ingresar Numero de Identidad para poder realizar la Busqueda", vbExclamation, sMensaje
    Exit Sub
  End If
  
  If VerConexionInfoRestCentral = "ON" Then
    If Calcular("select count(tidentidad) as Codigo from TDELIVERY where tidentidad='" & Trim(txtNumeroIdentidad.Text) & "'", Conex) > 0 Then
       MsgBox "Cliente si Existe en Base Central", vbExclamation, sMensaje
       Isql = "select isnull(tTipoCliente,'') as tTipoCliente , isnull(TipoCliente,'')as  TipoCliente, tTelefono, isnull(tTipoIdentidad,'')as tTipoIdentidad,ISNULL(TipoIdentidad,'')AS TipoIdentidad , tApellido,tNombre,tDireccion,tReferencia,isnull(tDistrito,'')as tDistrito,isnull(Distrito,'') as Distrito,tZona,isnull(Zona,'')as Zona,tObservacion,tEmail,fNacimiento,lActivo From vDelivery Where tidentidad='" & Trim(txtNumeroIdentidad.Text) & "'"
        
       Set xrsParientes = Lib.OpenRecordset(Isql, Conex)
        sTipoIdentidad = xrsParientes!tTipoIdentidad
        
        txtTipoCliente.Text = xrsParientes!tTipoCliente
        txtTipoClienteDetalle.Text = xrsParientes!TipoCliente
        txtTipoIdentidad.Text = xrsParientes!TipoIdentidad
        txtTelefono.Text = xrsParientes!tTelefono
        txtApellido.Text = xrsParientes!tApellido
        txtNombre.Text = xrsParientes!tNombre
        txtDireccion.Text = xrsParientes!tDireccion
        txtReferencia.Text = xrsParientes!tReferencia
        txtObservacion.Text = xrsParientes!tObservacion
        txtEmail.Text = xrsParientes!tEmail
        
        txtDistritoCodigo.Text = xrsParientes!tDistrito
        txtDistritoDetalle.Text = xrsParientes!Distrito
           
        txtZonaCodigo.Text = xrsParientes!tZona
        txtZonaDetalle.Text = xrsParientes!Zona
        
        
    Else
        MsgBox "Cliente No Existe", vbExclamation, sMensaje
    End If

  End If
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    frmKeyBoard.txtResultado = txtNumero.Text
    frmKeyBoard.Caption = "Número de Tarjeta"
    frmKeyBoard.Show vbModal
    txtNumero.Text = IIf(wEnter = True, sDescrip, txtNumero.Text)
End Sub

Private Sub cmdBuscarDistrito_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vDistrito where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                                                    
    frmBusquedaRapida.nPredeterm = 1
    sTipo = ""
    sTemp = txtDistritoDetalle.Text
    
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sDistrito = sCodigo
       RsDistrito.MoveFirst
       RsDistrito.Find ("Codigo = '" & sDistrito & "'")
       If Not RsDistrito.EOF Then
          txtDistritoCodigo.Text = IIf(IsNull(RsDistrito!codigo), "", RsDistrito!codigo)
          txtDistritoDetalle.Text = IIf(IsNull(RsDistrito!Descripcion), "", RsDistrito!Descripcion)
       Else
          sDistrito = ""
          txtDistritoCodigo.Text = ""
          txtDistritoDetalle.Text = ""
       End If
       SendKeys "{Tab}"
    End If
End Sub

Private Sub cmdBuscarTarjeta_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    sTemp = txtTarjeta.Text
    Isql = "select tCodigoTarjeta as Codigo, tDetallado as Descripcion from TTARJETACREDITO where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 2200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 6000, 0, 0, "")
 
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter = True Then
       txtTarjeta.Text = sDescrip
       sTarjeta = sCodigo
    End If
End Sub

Private Sub cmdBuscarZona_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    Dim xCodigo As String
    Isql = "select * from vZona where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 2000, 0, 0, "", _
                                                    "Resumen", 2, "tresumido", 5000, 0, 0, "")     'CD
    sTemp = txtZonaDetalle.Text
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sZona = sCodigo
       RsZona.Requery
       RsZona.MoveFirst
       RsZona.Find ("Codigo = '" & sCodigo & "'")
       If Not RsZona.EOF Then
          sZona = sCodigo
          txtZonaCodigo.Text = IIf(IsNull(RsZona!codigo), "", RsZona!codigo)
          txtZonaDetalle.Text = IIf(IsNull(RsZona!Descripcion), "", RsZona!Descripcion)
       Else
          sZona = ""
          txtZonaCodigo.Text = ""
          txtZonaDetalle.Text = ""
       End If
       SendKeys "{Tab}"
    End If
End Sub

Private Sub cmdBuscaTipoCliente_Click()
   If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vTipoClienteFrecuente where codigo <> '00' and lActivo = 1 Order by codigo"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    'CESAR NORKYS
    sTemp = txtTipoClienteDetalle.Text
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sTipoCliente = sCodigo
       RsTipoCliente.MoveFirst
       RsTipoCliente.Find ("Codigo = '" & sTipoCliente & "'")
       If Not RsTipoCliente.EOF Then
          txtTipoCliente.Text = IIf(IsNull(RsTipoCliente!codigo), "", RsTipoCliente!codigo)
          txtTipoClienteDetalle.Text = IIf(IsNull(RsTipoCliente!Descripcion), "", RsTipoCliente!Descripcion)
       Else
          sTipoCliente = ""
          txtTipoCliente.Text = ""
          txtTipoClienteDetalle.Text = ""
       End If
    End If
End Sub

Private Sub cmdEstado_Click()
   If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vEstadoFrecuente where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    sTemp = txtEstadoDescripcion.Text
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sEstado = sCodigo
       RsEstado.MoveFirst
       RsEstado.Find ("Codigo = '" & sEstado & "'")
       If Not RsEstado.EOF Then
          txtEstadoCodigo.Text = IIf(IsNull(RsEstado!codigo), "", RsEstado!codigo)
          txtEstadoDescripcion.Text = IIf(IsNull(RsEstado!Descripcion), "", RsEstado!Descripcion)
       Else
          sEstado = ""
          txtEstadoCodigo.Text = ""
          txtEstadoDescripcion.Text = ""
       End If
    End If
End Sub

Private Sub cmdLinea_Click()
    sTipo = ""
    frmNumPad.Show vbModal
    txtLinea.Text = Format(IIf(wEnter = True, sDescrip, txtLinea.Text), "###,###,###,##0.00")
End Sub

Private Sub cmdNumeroIdentidad_Click()
    frmKeyBoard.txtResultado = txtNumeroIdentidad.Text
    frmKeyBoard.Caption = "Número de Identidad"
    frmKeyBoard.Show vbModal
    txtNumeroIdentidad.Text = IIf(wEnter = True, sDescrip, txtNumeroIdentidad.Text)
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
          Case Is = 0 'Apellido
               frmKeyBoard.txtResultado = txtApellido.Text
               frmKeyBoard.Caption = "Apellidos"
               frmKeyBoard.Show vbModal
               txtApellido.Text = IIf(wEnter, sDescrip, txtApellido.Text)
          
          Case Is = 1 ' Nombres
               frmKeyBoard.txtResultado = txtNombre.Text
               frmKeyBoard.Caption = "Nombres"
               frmKeyBoard.Show vbModal
               txtNombre.Text = IIf(wEnter, sDescrip, txtNombre.Text)
          
          Case Is = 2 ' Direccion
               frmKeyBoard.txtResultado = txtDireccion.Text
               frmKeyBoard.Caption = "Dirección"
               frmKeyBoard.Show vbModal
               txtDireccion.Text = IIf(wEnter, sDescrip, txtDireccion.Text)
               
          Case Is = 3 ' Telefono
               frmKeyBoard.txtResultado = txtTelefono.Text
               frmKeyBoard.Caption = "Teléfono"
               frmKeyBoard.Show vbModal
               txtTelefono.Text = IIf(wEnter, sDescrip, txtTelefono.Text)
               
          Case Is = 4 ' Referencia
               frmKeyBoard.txtResultado = txtReferencia.Text
               frmKeyBoard.Caption = "Referencia"
               frmKeyBoard.Show vbModal
               txtReferencia.Text = IIf(wEnter, sDescrip, txtReferencia.Text)
               
          Case Is = 5 ' Grabar
               Dim nCorrela As String
               Dim xnCorrela As String
               'Chequea Datos
               If txtTelefono.Text = "" Then MsgBox "Ingrese el Telefono O Id", vbExclamation, sMensaje: Exit Sub
               If txtTipoCliente.Text = "" Then MsgBox ("Ingrese Tipo de Cliente"), vbExclamation, sMensaje: Exit Sub
               If Len(Trim(txtTelefono.Text)) = 8 And Trim(txtNumeroIdentidad.Text) = "" Then
                     txtNumeroIdentidad.Text = Trim(txtTelefono.Text)
               End If
               
               If frmBusquedaDelivery.wAdiciona Then
                  
                If VerConexionInfoRestCentral = "ON" Then
                    If Calcular("select count(tIdentidad) as Codigo from TDELIVERY where tIdentidad='" & Trim(txtNumeroIdentidad.Text) & "'", Cn) > 0 Then
                       MsgBox "DNI Existente", vbExclamation, sMensaje
                       Exit Sub
                    End If
                Else
                    If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                       MsgBox "Teléfono o Id Existente", vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
                
                '------------------------
                '------------------------ SI ESTA ACTIVO PARA GRABAR CLIENTE
                '=====================GCAA 27/02/2020 GRABAR CLIENTE
                   'Obtiene el Numero de Orden
                  
                  If VerConexionInfoRestCentral = "ON" Then
                        If Calcular("select count(tIdentidad) as Codigo from TDELIVERY where tIdentidad='" & Trim(txtNumeroIdentidad.Text) & "'", Conex) > 0 Then
                           Else
                                xnCorrela = Calcular("select max(tCodigodelivery) as Codigo from TDELIVERY", Conex)
                                If Mid(xnCorrela, 1, 1) = "0" Or Mid(xnCorrela, 1, 1) = "1" Or Mid(xnCorrela, 1, 1) = "2" Or Mid(xnCorrela, 1, 1) = "3" Or Mid(xnCorrela, 1, 1) = "4" Or Mid(xnCorrela, 1, 1) = "5" Or Mid(xnCorrela, 1, 1) = "6" Or Mid(xnCorrela, 1, 1) = "7" Or Mid(xnCorrela, 1, 1) = "8" Or Mid(xnCorrela, 1, 1) = "9" Then
                                          If IsNull(xnCorrela) Or xnCorrela = "" Then
                                              xsCliente = "0000001"
                                          Else
                                              xsCliente = Lib.Correlativo(xnCorrela, 7)
                                          End If
                                 Else
                                          If IsNull(xnCorrela) Or xnCorrela = "" Then
                                              xsCliente = "0000001"
                                          Else
                                              If Mid(xnCorrela, 2, 6) = "999999" Then
                                              xsCliente = Chr(Asc(Mid(xnCorrela, 1, 1)) + 1) + "000001"
                                              Else
                                              xsCliente = Mid(xnCorrela, 1, 1) + Lib.Correlativo(Mid(xnCorrela, 2, 6), 6)
                                              End If
                                 
                                          End If
                                 End If
                                 
                                   Isql = "insert into TDELIVERY( " & _
                                    "tCodigoDelivery, tTipoCliente, tApellido, tNombre, tDireccion, tTelefono, tReferencia, tObservacion, tEmail, ndescuento, fNacimiento, tZona, lActivo, lPuntos, nAcumulado, nUtilizado, nDisponible, tCodigoTarjeta, tNumeroTarjeta, tDistrito, lClienteCtaCte, nLinea, tTipoCtaCte, tSubTipoCtaCte, tEstadoFrecuente, tTipoIdentidad,  tIdentidad, tUsuario, fRegistro) " & _
                                    "values ('" & xsCliente & "', " & _
                                            " '" & sTipoCliente & "', " & _
                                            " '" & txtApellido.Text & "', " & _
                                            " '" & txtNombre.Text & "', " & _
                                            " '" & txtDireccion.Text & "', " & _
                                            " '" & Trim(txtTelefono.Text) & "', " & _
                                            " '" & txtReferencia.Text & "', " & _
                                            " '" & txtObservacion.Text & "', " & _
                                            " '" & txtEmail.Text & "', " & _
                                            " '" & IIf(IsNull(txtDescuento.Text), 0, txtDescuento.Text) & "'," & _
                                            IIf(dtpfecha.Enabled, "'" & Format(dtpfecha.value, "yyyy/MM/dd") & "'", "Null") & ", " & _
                                            "'" & Trim(sZona) & "', 1, " & IIf(txtPuntos.Text = "Si", 1, 0) & ", 0, 0, 0, " & _
                                            " '" & sTarjeta & "', '" & txtNumero.Text & "', '" & sDistrito & "' , " & chkCtaCte.value & ", " & CDbl(IIf(txtLinea.Text = "", 0, txtLinea.Text)) & ", '" & sTipoCuentaCorriente & "', '" & sTipoSubCuentaCorriente & "', '" & sEstado & "', '" & sTipoIdentidad & "', '" & txtNumeroIdentidad.Text & "', " & _
                                            " '" & sUsuario & "', " & _
                                            " getdate())"
                                Conex.Execute Isql
                            End If
                    End If ' TERMINAR LA CONDICION SI ESTA PRENDIDO EL CENTRAL INFOREST
                       
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigodelivery) as Codigo from TDELIVERY", Cn)
                  If Mid(nCorrela, 1, 1) = "0" Or Mid(nCorrela, 1, 1) = "1" Or Mid(nCorrela, 1, 1) = "2" Or Mid(nCorrela, 1, 1) = "3" Or Mid(nCorrela, 1, 1) = "4" Or Mid(nCorrela, 1, 1) = "5" Or Mid(nCorrela, 1, 1) = "6" Or Mid(nCorrela, 1, 1) = "7" Or Mid(nCorrela, 1, 1) = "8" Or Mid(nCorrela, 1, 1) = "9" Then
                            If IsNull(nCorrela) Or nCorrela = "" Then
                                sCliente = "0000001"
                            Else
                                sCliente = Lib.Correlativo(nCorrela, 7)
                            End If
                   Else
                   ' MsgBox Asc(Mid(nCorrela, 1, 1))
                            If IsNull(nCorrela) Or nCorrela = "" Then
                                sCliente = "0000001"
                            Else
                            If Mid(nCorrela, 2, 6) = "999999" Then
                            sCliente = Chr(Asc(Mid(nCorrela, 1, 1)) + 1) + "0000001"
                            Else
                                sCliente = Mid(nCorrela, 1, 1) + Lib.Correlativo(Mid(nCorrela, 2, 6), 6)
                                End If
                            End If
                   End If

                  Isql = "insert into TDELIVERY( " & _
                         "tCodigoDelivery, tTipoCliente, tApellido, tNombre, tDireccion, tTelefono, tReferencia, tObservacion, tEmail, ndescuento, fNacimiento, tZona, lActivo, lPuntos, nAcumulado, nUtilizado, nDisponible, tCodigoTarjeta, tNumeroTarjeta, tDistrito, lClienteCtaCte, nLinea, tTipoCtaCte, tSubTipoCtaCte, tEstadoFrecuente, tTipoIdentidad,  tIdentidad, tUsuario, fRegistro) " & _
                         "values ('" & sCliente & "', " & _
                                 " '" & sTipoCliente & "', " & _
                                 " '" & txtApellido.Text & "', " & _
                                 " '" & txtNombre.Text & "', " & _
                                 " '" & txtDireccion.Text & "', " & _
                                 " '" & Trim(txtTelefono.Text) & "', " & _
                                 " '" & txtReferencia.Text & "', " & _
                                 " '" & txtObservacion.Text & "', " & _
                                 " '" & txtEmail.Text & "', " & _
                                 " '" & IIf(IsNull(txtDescuento.Text), 0, txtDescuento.Text) & "'," & _
                                 IIf(dtpfecha.Enabled, "'" & Format(dtpfecha.value, "yyyy/MM/dd") & "'", "Null") & ", " & _
                                 "'" & Trim(sZona) & "', 1, " & IIf(txtPuntos.Text = "Si", 1, 0) & ", 0, 0, 0, " & _
                                 " '" & sTarjeta & "', '" & txtNumero.Text & "', '" & sDistrito & "' , " & chkCtaCte.value & ", " & CDbl(IIf(txtLinea.Text = "", 0, txtLinea.Text)) & ", '" & sTipoCuentaCorriente & "', '" & sTipoSubCuentaCorriente & "', '" & sEstado & "', '" & sTipoIdentidad & "', '" & txtNumeroIdentidad.Text & "', " & _
                                 " '" & sUsuario & "', " & _
                                 " getdate())"
                  Cn.Execute Isql
               Else
                  If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tCodigoDelivery<>'" & sCliente & "' and tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                     MsgBox "Teléfono o Id Existente", vbExclamation, sMensaje
                     Exit Sub
                  End If

                  Isql = "Update TDELIVERY SET " & _
                         "tApellido='" & txtApellido.Text & "', " & _
                         "tNombre='" & txtNombre.Text & "', " & _
                         "tDireccion='" & txtDireccion.Text & "', " & _
                         "tTelefono='" & Trim(txtTelefono.Text) & "', " & _
                         "tReferencia='" & txtReferencia.Text & "', " & _
                         "tObservacion='" & txtObservacion.Text & "', " & _
                         "tEmail='" & txtEmail.Text & "', " & _
                         "fNacimiento= " & IIf(dtpfecha.Enabled, "'" & Format(dtpfecha.value, "yyyy/MM/dd") & "'", "Null") & ", " & _
                         "tZona='" & Trim(sZona) & "', " & _
                         "tNumeroTarjeta='" & txtNumero.Text & "', tCodigoTarjeta='" & sTarjeta & "', " & _
                         "lPuntos=" & IIf(txtPuntos.Text = "Si", 1, 0) & ", " & _
                         "nDescuento='" & txtDescuento.Text & "' , lClienteCtaCte = '" & chkCtaCte.value & "', tDistrito='" & sDistrito & "', lReplica=1, nlinea=" & CDbl(IIf(txtLinea.Text = "", 0, txtLinea.Text)) & ", tTipoCtaCte='" & sTipoCuentaCorriente & "', tSubTipoCtaCte='" & sTipoSubCuentaCorriente & "',tEstadoFrecuente='" & sEstado & "', tTipoIdentidad='" & sTipoIdentidad & "',tIdentidad='" & Me.txtNumeroIdentidad.Text & "' " & _
                         "where tCodigoDelivery ='" & sCliente & "'"
                   Cn.Execute Isql
               End If
               
               'Cn.Execute "Update TDELIVERY SET lNuevoModificado=1 where tCodigoDelivery ='" & sCliente & "'"
               wEnter = True
               sCodigo = sCliente
               Unload Me
          
          Case Is = 6 ' Salir
               wEnter = False
               Unload Me
               
          Case Is = 7 ' Observacion
               frmKeyBoard.txtResultado = txtObservacion.Text
               frmKeyBoard.Caption = "Observación"
               frmKeyBoard.Show vbModal
               txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
          
          Case Is = 8 ' Descuento
               If Supervisor("10") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               sTipo = ""
               frmNumPad.txtResultado = txtDescuento.Text
               frmNumPad.Caption = "Descuento"
               frmNumPad.Show vbModal
               txtDescuento.Text = Format(IIf(wEnter, sDescrip, txtDescuento.Text), "##0.00")
            
          Case Is = 9 ' Tiendas
               If sCliente = "" Then
                  MsgBox "Agregue primero al Cliente", vbCritical + vbOKOnly, sMensaje
                  Exit Sub
               End If
               sCodigo = sCliente
               sDescrip = txtApellido.Text
               sTemp = txtApellido.Text
               'frmTienda.Show vbModal
            
          Case Is = 10 ' Canje por puntos
               txtPuntos.Text = IIf(txtPuntos.Text = "Si", "No", "Si")
                              
          Case Is = 11 ' Fecha Nacimiento
               dtpfecha.Enabled = Not dtpfecha.Enabled
               
          Case Is = 12 ' Email
               frmKeyBoard.txtResultado = txtEmail.Text
               frmKeyBoard.Caption = "E-mail"
               frmKeyBoard.Show vbModal
               txtEmail.Text = IIf(wEnter, sDescrip, txtEmail.Text)
   End Select

End Sub

Private Sub cmdTipoIdentidad_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vTipoIdentidad where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                                                    
    frmBusquedaRapida.nPredeterm = 1
    sTemp = txtTipoIdentidad.Text
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sTipoIdentidad = sCodigo
       RsTipoIdentidad.MoveFirst
       RsTipoIdentidad.Find ("Codigo = '" & sTipoIdentidad & "'")
       If Not RsTipoIdentidad.EOF Then
          txtTipoIdentidad.Text = IIf(IsNull(RsTipoIdentidad!Descripcion), "", RsTipoIdentidad!Descripcion)
       Else
          sTipoIdentidad = ""
          txtTipoIdentidad.Text = ""
       End If
    End If
End Sub

Private Sub cmdTipoSubCuentaCorriente_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vSubTipoCtaCte where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                                                    
    frmBusquedaRapida.nPredeterm = 1
    sTemp = txtTipoSubCuentaCorrienteDescripcion
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sTipoSubCuentaCorriente = sCodigo
       RsTipoSubCuentaCorriente.MoveFirst
       RsTipoSubCuentaCorriente.Find ("Codigo = '" & sTipoSubCuentaCorriente & "'")
       If Not RsTipoSubCuentaCorriente.EOF Then
          txtTipoSubCuentaCorrienteCodigo.Text = IIf(IsNull(RsTipoSubCuentaCorriente!codigo), "", RsTipoSubCuentaCorriente!codigo)
          txtTipoSubCuentaCorrienteDescripcion.Text = IIf(IsNull(RsTipoSubCuentaCorriente!Descripcion), "", RsTipoSubCuentaCorriente!Descripcion)
       Else
          sTipoSubCuentaCorriente = ""
          txtTipoSubCuentaCorrienteCodigo.Text = ""
          txtTipoSubCuentaCorrienteDescripcion.Text = ""
       End If
    End If

End Sub

Private Sub cmdTipoCuentaCorriente_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
        
    Isql = "select * from vTipoCtaCte where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                                                    
    frmBusquedaRapida.nPredeterm = 1
    sTemp = txtTipoCuentaCorrienteDetalle.Text
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sTipoCuentaCorriente = sCodigo
       RsTipoCuentaCorriente.MoveFirst
       RsTipoCuentaCorriente.Find ("Codigo = '" & sTipoCuentaCorriente & "'")
       If Not RsTipoCuentaCorriente.EOF Then
          txtTipoCuentaCorrienteCodigo.Text = IIf(IsNull(RsTipoCuentaCorriente!codigo), "", RsTipoCuentaCorriente!codigo)
          txtTipoCuentaCorrienteDetalle.Text = IIf(IsNull(RsTipoCuentaCorriente!Descripcion), "", RsTipoCuentaCorriente!Descripcion)
       Else
          sTipoCuentaCorriente = ""
          txtTipoCuentaCorrienteCodigo.Text = ""
          txtTipoCuentaCorrienteDetalle.Text = ""
       End If
    End If
End Sub





Private Sub Form_Load()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   sTipo = ""
   Centrar Me
   Set RsZona = Lib.OpenRecordset("select * from vZona where lActivo = 1", Cn)
   Set RsDistrito = Lib.OpenRecordset("select * from vDistrito where lActivo = 1", Cn)
   Set RsTipoIdentidad = Lib.OpenRecordset("select * from vTipoIdentidad where lActivo = 1", Cn)
   Set RsTipoCuentaCorriente = Lib.OpenRecordset("select * from vTipoCtaCte where lActivo = 1", Cn)
   Set RsTipoSubCuentaCorriente = Lib.OpenRecordset("select * from vSubTipoCTaCte where lActivo = 1", Cn)
   Set RsEstado = Lib.OpenRecordset("select * from vEstadoFrecuente where lActivo = 1", Cn)
   Set RsTipoCliente = Lib.OpenRecordset("select * from vTipoClienteFrecuente where codigo <> '00' and  lActivo = 1", Cn)
   
   sZona = ""
   sDistrito = ""
   sTipoIdentidad = ""
   sTipoCuentaCorriente = ""
   sTipoSubCuentaCorriente = ""
         
   If frmBusquedaDelivery.wAdiciona Then
       Blanquear Me
       txtPuntos.Text = "Si"
       txtDescuento.Text = "0.00"
       sEstado = "01"
       txtEstadoCodigo.Text = sEstado
       txtEstadoDescripcion.Text = Calcular("select Descripcion as Codigo from vEstadoFrecuente where Codigo='" & sEstado & "'", Cn)
   
       sTipoCliente = Calcular("select Codigo as Codigo from vTipoClienteFrecuente where lActivo=1 and nValor=1", Cn)
       If sTipoCliente = "0" Then
          sTipoCliente = "" '"00"
          txtTipoClienteDetalle.Text = "" ' "SIN TIPO"
       Else
          txtTipoClienteDetalle.Text = Calcular("select Descripcion as Codigo from vTipoClienteFrecuente where lActivo=1 and nValor=1", Cn)
       End If
       txtTipoCliente.Text = sTipoCliente
   Else
       sCliente = IIf(frmBusquedaDelivery.RsGrilla.EOF = True, "", frmBusquedaDelivery.RsGrilla!codigo)
       Mostrar
   End If
   
   SSTab1.Tab = 0
   
End Sub

Sub Mostrar()
    Set RsDelivery = Lib.OpenRecordset("SELECT * from vDelivery where Codigo ='" & frmBusquedaDelivery.RsGrilla!codigo & "'", Cn)
    Dim i As Integer
    With RsDelivery
        txtApellido.Text = IIf(IsNull(!tApellido), "", !tApellido)
        txtNombre.Text = IIf(IsNull(!tNombre), "", !tNombre)
        txtDireccion.Text = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtTelefono.Text = IIf(IsNull(!tTelefono), "", !tTelefono)
        txtReferencia.Text = IIf(IsNull(!tReferencia), "", !tReferencia)
        txtObservacion.Text = IIf(IsNull(!tObservacion), "", !tObservacion)
        txtEmail.Text = IIf(IsNull(!tEmail), "", !tEmail)
        txtNumero.Text = IIf(IsNull(!tNumeroTarjeta), "", !tNumeroTarjeta)
        txtDescuento.Text = Format(IIf(IsNull(!nDescuento), 0, !nDescuento), "##0.00")
        txtPuntos.Text = IIf(!lPuntos, "Si", "No")
        dtpfecha.Enabled = IIf(IsNull(!fNacimiento), False, True)
        dtpfecha.value = IIf(IsNull(!fNacimiento), FechaServidor(), !fNacimiento)
        
        sTarjeta = IIf(IsNull(!tCodigoTarjeta), "", !tCodigoTarjeta)
        txtTarjeta.Text = IIf(IsNull(!TarjetaCredito), "", !TarjetaCredito)
        txtNumero.Text = IIf(IsNull(!tNumeroTarjeta), "", !tNumeroTarjeta)
        sDistrito = IIf(IsNull(!tDistrito), "", !tDistrito)
        txtDistritoCodigo.Text = sDistrito
        txtDistritoDetalle.Text = IIf(IsNull(!Distrito), "", !Distrito)
        sZona = IIf(IsNull(!tZona), "", !tZona)
        txtZonaCodigo.Text = sZona
        txtZonaDetalle.Text = IIf(IsNull(!Zona), "", !Zona)
        sEstado = IIf(IsNull(!tEstadoFrecuente), "", !tEstadoFrecuente)
        txtEstadoCodigo.Text = sEstado
        txtEstadoDescripcion.Text = IIf(IsNull(!EstadoFrecuente), "", !EstadoFrecuente)
        txtTipoCliente.Text = IIf(IsNull(!tTipoCliente), "", !tTipoCliente)
        txtTipoClienteDetalle.Text = IIf(IsNull(!TipoCliente), "", !TipoCliente)
        
        chkCtaCte.value = IIf(!lClienteCtaCte = True, 1, 0)
        If chkCtaCte.value = 1 Then
            sTipoIdentidad = IIf(IsNull(!tTipoIdentidad), "", !tTipoIdentidad)
            txtTipoIdentidad.Text = IIf(IsNull(!TipoIdentidad), "", !TipoIdentidad)
            txtNumeroIdentidad.Text = IIf(IsNull(!tIdentidad), "", !tIdentidad)
            sTipoCuentaCorriente = IIf(IsNull(!tTipoCtaCte), "", !tTipoCtaCte)
            txtTipoCuentaCorrienteCodigo.Text = sTipoCuentaCorriente
            txtTipoCuentaCorrienteDetalle.Text = IIf(IsNull(!TipoCtaCte), "", !TipoCtaCte)
            sTipoSubCuentaCorriente = IIf(IsNull(!tSubTipoCtaCte), "", !tSubTipoCtaCte)
            txtTipoSubCuentaCorrienteCodigo.Text = sTipoSubCuentaCorriente
            txtTipoSubCuentaCorrienteDescripcion.Text = IIf(IsNull(!SubTipoCtaCte), "", !SubTipoCtaCte)
            txtLinea.Text = Format(IIf(IsNull(!nLinea), "", !nLinea), "###,###,##0.00")
        Else
            txtLinea.Text = "0.00"
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmNuevoDelivery = Nothing
End Sub



Private Sub txtApellido_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtApellido_LostFocus()
   Call ValidaStr(txtApellido)
   txtApellido.Text = UCase(txtApellido.Text)
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtDireccion_LostFocus()
   Call ValidaStr(txtDireccion)
   txtDireccion.Text = UCase(txtDireccion.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtNombre_LostFocus()
   Call ValidaStr(txtNombre)
   txtNombre.Text = UCase(txtNombre.Text)
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtObservacion_LostFocus()
   Call ValidaStr(txtObservacion)
   txtObservacion.Text = UCase(txtObservacion.Text)
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtReferencia_LostFocus()
   Call ValidaStr(txtReferencia)
   txtReferencia.Text = UCase(txtReferencia.Text)
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtTelefono_LostFocus()
   Call ValidaStr(txtTelefono)
End Sub


