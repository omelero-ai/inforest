VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGrupoAccesoDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo de Usuarios"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkModulo 
      Caption         =   "Autorizaciones"
      Height          =   195
      Index           =   5
      Left            =   8400
      TabIndex        =   79
      Top             =   840
      Width           =   1395
   End
   Begin VB.CheckBox chkModulo 
      Caption         =   "Auditoría"
      Height          =   195
      Index           =   4
      Left            =   6480
      TabIndex        =   13
      Top             =   840
      Width           =   1395
   End
   Begin VB.CheckBox chkModulo 
      Caption         =   "Consultas y Reportes"
      Height          =   195
      Index           =   3
      Left            =   6480
      TabIndex        =   15
      Top             =   1080
      Width           =   2280
   End
   Begin VB.CheckBox chkModulo 
      Caption         =   "Administración y Configuración"
      Height          =   195
      Index           =   2
      Left            =   3420
      TabIndex        =   12
      Top             =   840
      Width           =   2715
   End
   Begin VB.CheckBox chkActivo 
      Caption         =   "Activo :"
      Height          =   240
      Left            =   8400
      TabIndex        =   11
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkModulo 
      Caption         =   "Punto de Venta"
      Height          =   195
      Index           =   1
      Left            =   3420
      TabIndex        =   14
      Top             =   1080
      Width           =   1965
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9960
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8565
      Width           =   10020
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
         Index           =   3
         Left            =   8670
         Picture         =   "frmGrupoAccesoDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Index           =   2
         Left            =   7500
         Picture         =   "frmGrupoAccesoDetalle.frx":00F2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   5160
         Picture         =   "frmGrupoAccesoDetalle.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Index           =   1
         Left            =   6330
         Picture         =   "frmGrupoAccesoDetalle.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4950
         TabIndex        =   44
         Top             =   60
         Width           =   5010
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmGrupoAccesoDetalle.frx":0C58
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmGrupoAccesoDetalle.frx":119A
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmGrupoAccesoDetalle.frx":16DC
            Style           =   1  'Graphical
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4470
            Picture         =   "frmGrupoAccesoDetalle.frx":1C1E
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmGrupoAccesoDetalle.frx":2160
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3510
            Picture         =   "frmGrupoAccesoDetalle.frx":26A2
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   45
            Top             =   165
            Width           =   1980
         End
      End
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   90
      Width           =   1170
   End
   Begin VB.TextBox txtDetallado 
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
      Left            =   3420
      MaxLength       =   15
      TabIndex        =   10
      Text            =   " "
      Top             =   465
      Width           =   4080
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   7305
      Left            =   2340
      TabIndex        =   33
      Top             =   1320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   12885
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Autorizaciones"
      TabPicture(0)   =   "frmGrupoAccesoDetalle.frx":2BE4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Punto de Venta"
      TabPicture(1)   =   "frmGrupoAccesoDetalle.frx":2C00
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvPuntoVenta"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Administración"
      TabPicture(2)   =   "frmGrupoAccesoDetalle.frx":2C1C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tvAdministracion"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Reportes y Consultas"
      TabPicture(3)   =   "frmGrupoAccesoDetalle.frx":2C38
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tvConsulta"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   " Acceso Rapido "
         Height          =   4380
         Left            =   -68970
         TabIndex        =   63
         Top             =   360
         Width           =   1455
         Begin VB.CheckBox chkBotonC7 
            Caption         =   "Propinas"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   3735
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC6 
            Caption         =   "Paloteo de Producción"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   3155
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC5 
            Caption         =   "Registro de Ventas"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   2578
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC4 
            Caption         =   "Liquidación de Cajero"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   2001
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC3 
            Caption         =   "Correlativo Documentos"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   1424
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC2 
            Caption         =   "Correlativo de Pedidos"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   847
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonC1 
            Caption         =   "Vistas de Mesas"
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
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   270
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Acceso Rapido "
         Height          =   4380
         Left            =   -70095
         TabIndex        =   56
         Top             =   360
         Width           =   2580
         Begin VB.CheckBox chkBotonP14 
            Caption         =   "Delivery Entregados"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   3735
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP8 
            Caption         =   "Correlativo de Pedidos"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   270
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP9 
            Caption         =   "Correlativo Documentos"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   840
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP10 
            Caption         =   "Cuentas Corrietes"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1425
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP11 
            Caption         =   "Cuentas por Cobrar"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   1995
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP12 
            Caption         =   "Carta de Productos"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   2580
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP13 
            Caption         =   "Delivery en Tránsito"
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
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   3150
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP7 
            Caption         =   "Reservas"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   3735
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP6 
            Caption         =   "Recibo de Ingresos"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   3155
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP5 
            Caption         =   "Recibo de Egresos"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2578
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP4 
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
            Height          =   555
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2001
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP3 
            Caption         =   "Cierre"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1424
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP2 
            Caption         =   "Punto de Venta"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   847
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonP1 
            Caption         =   "Apertura"
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   270
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Acceso Rapido "
         Height          =   4425
         Left            =   -68970
         TabIndex        =   48
         Top             =   315
         Width           =   1455
         Begin VB.CheckBox chkBotonA1 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   315
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonA2 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   972
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonA3 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1629
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonA4 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2286
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonA5 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2943
            Width           =   1185
         End
         Begin VB.CheckBox chkBotonA6 
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
            Height          =   555
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   3600
            Width           =   1185
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
         Height          =   6900
         Left            =   135
         TabIndex        =   37
         Top             =   310
         Width           =   7305
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Autoriza emision de recibo de egreso."
            Height          =   255
            Index           =   34
            Left            =   3720
            TabIndex        =   92
            Top             =   1680
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Autoriza pedido sin mesa"
            Height          =   255
            Index           =   33
            Left            =   3720
            TabIndex        =   91
            Top             =   1440
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar canal de venta del pedido"
            Height          =   195
            Index           =   32
            Left            =   3720
            TabIndex        =   90
            Top             =   1200
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar cortesia del documento"
            Height          =   255
            Index           =   31
            Left            =   3720
            TabIndex        =   89
            Top             =   840
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anular documento, despues de haber pasado los días como maximo"
            Height          =   435
            Index           =   30
            Left            =   3720
            TabIndex        =   88
            Top             =   360
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Bloquear emision de nota de credito a documentos fuera del mes"
            Height          =   195
            Index           =   29
            Left            =   270
            TabIndex        =   87
            Top             =   6600
            Width           =   6855
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Autoriza otros - pagos."
            Height          =   315
            Index           =   28
            Left            =   3720
            TabIndex        =   86
            Top             =   120
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emitir Notas de Credito"
            Height          =   195
            Index           =   27
            Left            =   270
            TabIndex        =   85
            Top             =   6360
            Width           =   2295
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar Mesa y #PAX  ---- Solo Funciona con Activacion de  Cover."
            Height          =   195
            Index           =   26
            Left            =   270
            TabIndex        =   84
            Top             =   6120
            Width           =   5655
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Genera Reimpresión de Pedidos Cuentas Corrientes"
            Height          =   195
            Index           =   25
            Left            =   270
            TabIndex        =   83
            Top             =   5860
            Width           =   4455
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Autoriza envíos de productos a producción"
            Height          =   195
            Index           =   24
            Left            =   270
            TabIndex        =   82
            Top             =   5600
            Width           =   3855
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Revertir el Despacho de un Pedido, Entregar Pedido No Cancelado ( Central de Pedido )"
            Height          =   195
            Index           =   23
            Left            =   270
            TabIndex        =   81
            Top             =   5330
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Controlar Nivel de Usuarios"
            Height          =   795
            Index           =   22
            Left            =   5520
            TabIndex        =   80
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar el Tipo de Tarifa a usar en Planilla de Motorizados"
            Height          =   195
            Index           =   21
            Left            =   270
            TabIndex        =   38
            Top             =   5070
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Acceso al Menú Procesos del Módulo Administración Centralizada"
            Height          =   195
            Index           =   20
            Left            =   270
            TabIndex        =   36
            Top             =   4820
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Acceso al Menú Configuraciones del Módulo Administración Centralizada"
            Height          =   195
            Index           =   19
            Left            =   270
            TabIndex        =   35
            Top             =   4570
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Administrador en Control de Turnos"
            Height          =   195
            Index           =   17
            Left            =   270
            TabIndex        =   32
            Top             =   4110
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Controler en Control de Turnos"
            Height          =   195
            Index           =   18
            Left            =   270
            TabIndex        =   34
            Top             =   4350
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Permiso a Multi Cajero que puede ver todas las cuentas, documentos y pagos."
            Height          =   195
            Index           =   16
            Left            =   270
            TabIndex        =   31
            Top             =   3870
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Generar Transferencias e Importaciones"
            Height          =   195
            Index           =   15
            Left            =   270
            TabIndex        =   30
            Top             =   3615
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Impresión de Precuentas y Reimpresiones de Pedidos"
            Height          =   195
            Index           =   14
            Left            =   270
            TabIndex        =   29
            Top             =   3375
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Eliminar Pedidos"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   16
            Top             =   150
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Eliminar Productos del Pedido"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   17
            Top             =   405
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar precios del Producto"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   18
            Top             =   645
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emitir Cortesías"
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   19
            Top             =   900
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anular y Cambiar documentos del turno"
            Height          =   195
            Index           =   5
            Left            =   270
            TabIndex        =   20
            Top             =   1140
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anulación, Modificación, Cambio de Documentos y Recibos de turnos anteriores"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   25
            Top             =   2385
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anular recibos del Turno"
            Height          =   195
            Index           =   7
            Left            =   270
            TabIndex        =   21
            Top             =   1395
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emitir Cuentas por Cobrar"
            Height          =   195
            Index           =   8
            Left            =   270
            TabIndex        =   22
            Top             =   1635
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emitir Cuentas Corrientes"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   23
            Top             =   1890
            Width           =   3500
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Generar Descuentos y Recargos en Pedidos, Documentos y Pagos"
            Height          =   195
            Index           =   10
            Left            =   270
            TabIndex        =   24
            Top             =   2130
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Generar Cierre de Turnos"
            Height          =   195
            Index           =   11
            Left            =   270
            TabIndex        =   26
            Top             =   2625
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Generar Reimpresiones de Precuentas y Documentos"
            Height          =   195
            Index           =   12
            Left            =   270
            TabIndex        =   27
            Top             =   2880
            Width           =   6615
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emitir Facturación o Anulación de Cuentas Corrientes"
            Height          =   195
            Index           =   13
            Left            =   270
            TabIndex        =   28
            Top             =   3120
            Width           =   6615
         End
      End
      Begin MSComctlLib.TreeView tvConsulta 
         Height          =   4335
         Left            =   -74955
         TabIndex        =   47
         Top             =   360
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   7646
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvAdministracion 
         Height          =   4335
         Left            =   -74955
         TabIndex        =   55
         Top             =   360
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   7646
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvPuntoVenta 
         Height          =   4335
         Left            =   -74955
         TabIndex        =   39
         Top             =   360
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   7646
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.Image Image 
      Height          =   8505
      Left            =   0
      Picture         =   "frmGrupoAccesoDetalle.frx":2C54
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2280
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Módulos :"
      Height          =   195
      Index           =   2
      Left            =   2625
      TabIndex        =   46
      Top             =   780
      Width           =   810
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Código :"
      Height          =   195
      Index           =   0
      Left            =   2745
      TabIndex        =   42
      Top             =   135
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      Height          =   195
      Index           =   1
      Left            =   2385
      TabIndex        =   41
      Top             =   510
      Width           =   930
   End
End
Attribute VB_Name = "frmGrupoAccesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupoAcceso As ADODB.Recordset
Dim RsAcceso As ADODB.Recordset

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmGrupoAcceso.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmGrupoAcceso.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmGrupoAcceso.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmGrupoAcceso.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmGrupoAcceso.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmGrupoAcceso.grdGrilla
    End Select
    Asignar
    cmdTexto.Caption = "Registro " & frmGrupoAcceso.RsCabecera.AbsolutePosition & " de " & frmGrupoAcceso.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Dim xGrupo As String
   Dim xGrupoCombo As String
   Dim xlControlNivel As Boolean
   
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Iniciar
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                    
               If Sw Then
                  'Asignar El Campo de Codificación
                   nCorrela = Calcular("select max(tGrupoUsuario) as Codigo from TGrupoUsuario", Cn)
                   If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                   Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                   End If
                   Sw = False
                  
 
                   
                 sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOUSUARIO", "GRUPO DE USUARIOS", "01", sUsuario, sPasa, "", _
                   "TGRUPOUSUARIO", "Codigo Grupo Usuario", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), _
                   "lModulo01", "Flag Modulo Punto de Venta", IIf(chkModulo(1).Value, "Verdadero", "Falso"), "lModulo02", "Flag Modulo Administracion", IIf(chkModulo(2).Value, "Verdadero", "Falso"), "lModulo03", "Flag Modulo Consultas", IIf(chkModulo(3).Value, "Verdadero", "Falso"), "lModulo04", "Flag Modulo Auditoria", IIf(chkModulo(4).Value, "Verdadero", "Falso"), _
                   "lOpcion01", "Flag Eliminar Pedidos", IIf(chkOpcion(1).Value, "Verdadero", "False"), "lOpcion02", "Flag Eliminar Productos de Pedido", IIf(chkOpcion(2).Value, "Verdadero", "Falso"), "lOpcion03", "Flag Modificar Precio Productos", IIf(chkOpcion(3).Value, "Verdadero", "Falso"), "lOpcion04", "Flag Emitir Cortesias", IIf(chkOpcion(4).Value, "Verdadero", "Falso"), "lOpcion05", "Flag Anular y Cambiar Documentos del Turno", IIf(chkOpcion(5).Value, "Verdadero", "Falso"), "lopcion06", "Flag Anulacion Cambio de Documento y Reb Turno Anterior", IIf(chkOpcion(6).Value, "Verdadero", "Falso"), "lOpcion07", "Flag Anular Recibos", IIf(chkOpcion(7).Value, "Verdadero", "Falso"), "lOpcion08", "Flag Emitir Cuentas Por Cobrar", IIf(chkOpcion(8).Value, "Verdadero", "Falso"), "lOpcion09", "Flag Emitir Cuentas Corrientes", IIf(chkOpcion(9).Value, "Verdadero", "Falso"), "lOpcion10", "Flag Descuentos y Recargos", IIf(chkOpcion(10).Value, "Verdadero", "Falso"), _
                   "lOpcion11", "Flag Generar Cierre de Turnos", IIf(chkOpcion(11).Value, "Verdadero", "Falso"), "lOpcion12", "Flag Genera Reimpresiones Documentos", IIf(chkOpcion(12).Value, "Verdadero", "Falso"), "lOpcion13", "Flag Emites Facturacion o Anulacion de Cta Cte", IIf(chkOpcion(13).Value, "Verdadero", "Falso"), "lOpcion14", "Flag Impresion Precuentas y Reimpresion de Pedido", IIf(chkOpcion(14).Value, "Verdadero", "Falso"), "lOpcion15", "Flag Generar Transferencia e Importaciones", IIf(chkOpcion(15).Value, "Verdadero", "Falso"), "lOpcion16", "Flag MultiCajero puede ser Todas Cuentas", IIf(chkOpcion(16).Value, "Verdadero", "Falso"), "lOpcion17", "Flag Administrador en Control Turnos", IIf(chkOpcion(17).Value, "Verdadero", "Falso"), "lOpcion18", "Flag Controler en Control de Turnos", IIf(chkOpcion(18).Value, "Verdadero", "Falso"), "lOpcion19", "Flag Menu Configuraciones ADM CEN", IIf(chkOpcion(19).Value, "Verdadero", "Falso"), _
                   "lOpcion20", "Flag Menu Procesos ADM CEN", IIf(chkOpcion(20).Value, "Verdadero", "Falso"), _
                   "lOpcion21", "Flag Modificar Tipo Tarifa Motorizados", IIf(chkOpcion(21).Value, "Verdadero", "Falso"), "lModulo05", "Flag Modulo Autorizaciones", IIf(chkModulo(5).Value, "Verdadero", "Falso"), "lOpcion22", "Flag Revierte Despacho", IIf(chkOpcion(23).Value, "Verdadero", "Falso"), "lOpcion29", "Flag bloquea emision de notas de credito", IIf(chkOpcion(29).Value, "Verdadero", "Falso"), "lOpcion30", "Anular Documento, despues de haber pasado los días como maximo", IIf(chkOpcion(30).Value, "Verdadero", "Falso"), "lOpcion31", "Modificar Cortesia del Pedido", IIf(chkOpcion(31).Value, "Verdadero", "Falso"), "lOpcion32", "Modificar Canal de Venta del Pedido", IIf(chkOpcion(32).Value, "Verdadero", "Falso"), "lOpcion33", "Activar Autorización de Pedido Sin Mesa", IIf(chkOpcion(33).Value, "Verdadero", "Falso"), "lOpcion34", "Autorizacion de recibo de egreso", IIf(chkOpcion(34).Value, "Verdadero", "Falso"))
                   
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If

                   'Cambiar el SQL
                   Isql = "insert into TGrupoUsuario( " & _
                          "tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01, lOpcion02, lOpcion03, lOpcion04, lOpcion05, lOpcion06, lOpcion07, lOpcion08, lOpcion09, lOpcion10, lOpcion11, lOpcion12, lOpcion13, lOpcion14, lOpcion15, lOpcion16,lOpcion17, lOpcion18, lActivo,lOpcion19, lOpcion20, LOPCION21, lModulo04, lModulo05, tNivel, lControlNivel, lopcion22, lOpcion24, lOpcion25, lOpcion26, lOpcion27, lOpcion28, lOpcion29, lOpcion30, lOpcion31, lOpcion32, lOpcion33, lOpcion34) " & _
                          "values ('" & txtCodigo.Text & "', '" & txtDetallado.Text & "', " & _
                                        chkModulo(1).Value & ", " & _
                                        chkModulo(2).Value & ", " & _
                                        chkModulo(3).Value & ", " & _
                                        chkOpcion(1).Value & ", " & _
                                        chkOpcion(2).Value & ", " & _
                                        chkOpcion(3).Value & ", " & _
                                        chkOpcion(4).Value & ", " & _
                                        chkOpcion(5).Value & ", " & _
                                        chkOpcion(6).Value & ", " & _
                                        chkOpcion(7).Value & ", " & _
                                        chkOpcion(8).Value & ", " & _
                                        chkOpcion(9).Value & ", " & _
                                        chkOpcion(10).Value & ", " & _
                                        chkOpcion(11).Value & ", " & _
                                        chkOpcion(12).Value & ", " & _
                                        chkOpcion(13).Value & ", " & _
                                        chkOpcion(14).Value & ", " & _
                                        chkOpcion(15).Value & ", " & _
                                        chkOpcion(16).Value & ", " & _
                                        chkOpcion(17).Value & ", " & _
                                        chkOpcion(18).Value & ", " & _
                                        chkActivo.Value & ", " & chkOpcion(19).Value & "," & chkOpcion(20).Value & ", " & chkOpcion(21).Value & "," & chkModulo(4).Value & "," & chkModulo(5).Value & ",'" & txtCodigo.Text & "'," & chkOpcion(22).Value & ", " & chkOpcion(23).Value & ", " & chkOpcion(24).Value & ", " & chkOpcion(25).Value & ", " & chkOpcion(26).Value & ", " & chkOpcion(27).Value & ", " & chkOpcion(28).Value & ", " & chkOpcion(29).Value & ", " & chkOpcion(30).Value & ",  " & chkOpcion(31).Value & ", " & chkOpcion(32).Value & ", " & chkOpcion(33).Value & ", " & chkOpcion(34).Value & ")"
                       Cn.Execute Isql
                       GrabaArbol
                       frmGrupoAcceso.RsCabecera.Sort = "tGrupoUsuario ASC"
                       frmGrupoAcceso.RsCabecera.Requery
                       frmGrupoAcceso.RsCabecera.MoveLast
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdTexto.Caption = "Registro " & IIf(frmGrupoAcceso.RsCabecera.RecordCount = 0, 0, frmGrupoAcceso.RsCabecera.AbsolutePosition) & " de " & frmGrupoAcceso.RsCabecera.RecordCount
               Else
               
                      
                 sPasa = txtCodigo.Text
                 
                  'CESAR VALIDACION MODIFICACION DE GRUPO
                  xGrupo = Calcular("Select tGrupoUsuario As Codigo From  TUSUARIO Where tResumido = '" & sUsuario & "' ", Cn)
                  xlControlNivel = Calcular("Select ISNULL(lControlUsuario,0) As Codigo From  TPARAMETRO", Cn)
                  xGrupoCombo = txtCodigo.Text
                  
                  
                  If xGrupo <> "00" Then
                        If xlControlNivel Then
                            If xGrupo > xGrupoCombo Then
                               MsgBox "No puede Modificar Grupos De Usuarios con un Nivel Superior", vbExclamation, sMensaje
                               Exit Sub
                            End If
                        End If
                  End If
                  '------------------
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOUSUARIO", "GRUPO DE USUARIOS", "02", sUsuario, sPasa, "", _
                   "TGRUPOUSUARIO", "Codigo Grupo Usuario", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), _
                   "lModulo01", "Flag Modulo Punto de Venta", IIf(chkModulo(1).Value, "Verdadero", "Falso"), "lModulo02", "Flag Modulo Administracion", IIf(chkModulo(2).Value, "Verdadero", "Falso"), "lModulo03", "Flag Modulo Consultas", IIf(chkModulo(3).Value, "Verdadero", "Falso"), "lModulo04", "Flag Modulo Auditoria", IIf(chkModulo(4).Value, "Verdadero", "Falso"), _
                   "lOpcion01", "Flag Eliminar Pedidos", IIf(chkOpcion(1).Value, "Verdadero", "False"), "lOpcion02", "Flag Eliminar Productos de Pedido", IIf(chkOpcion(2).Value, "Verdadero", "Falso"), "lOpcion03", "Flag Modificar Precio Productos", IIf(chkOpcion(3).Value, "Verdadero", "Falso"), "lOpcion04", "Flag Emitir Cortesias", IIf(chkOpcion(4).Value, "Verdadero", "Falso"), "lOpcion05", "Flag Anular y Cambiar Documentos del Turno", IIf(chkOpcion(5).Value, "Verdadero", "Falso"), "lopcion06", "Flag Anulacion Cambio de Documento y Reb Turno Anterior", IIf(chkOpcion(6).Value, "Verdadero", "Falso"), "lOpcion07", "Flag Anular Recibos", IIf(chkOpcion(7).Value, "Verdadero", "Falso"), "lOpcion08", "Flag Emitir Cuentas Por Cobrar", IIf(chkOpcion(8).Value, "Verdadero", "Falso"), "lOpcion09", "Flag Emitir Cuentas Corrientes", IIf(chkOpcion(9).Value, "Verdadero", "Falso"), "lOpcion10", "Flag Descuentos y Recargos", IIf(chkOpcion(10).Value, "Verdadero", "Falso"), _
                   "lOpcion11", "Flag Generar Cierre de Turnos", IIf(chkOpcion(11).Value, "Verdadero", "Falso"), "lOpcion12", "Flag Genera Reimpresiones Documentos", IIf(chkOpcion(12).Value, "Verdadero", "Falso"), "lOpcion13", "Flag Emites Facturacion o Anulacion de Cta Cte", IIf(chkOpcion(13).Value, "Verdadero", "Falso"), "lOpcion14", "Flag Impresion Precuentas y Reimpresion de Pedido", IIf(chkOpcion(14).Value, "Verdadero", "Falso"), "lOpcion15", "Flag Generar Transferencia e Importaciones", IIf(chkOpcion(15).Value, "Verdadero", "Falso"), "lOpcion16", "Flag MultiCajero puede ser Todas Cuentas", IIf(chkOpcion(16).Value, "Verdadero", "Falso"), "lOpcion17", "Flag Administrador en Control Turnos", IIf(chkOpcion(17).Value, "Verdadero", "Falso"), "lOpcion18", "Flag Controler en Control de Turnos", IIf(chkOpcion(18).Value, "Verdadero", "Falso"), "lOpcion19", "Flag Menu Configuraciones ADM CEN", IIf(chkOpcion(19).Value, "Verdadero", "Falso"), _
                   "lOpcion20", "Flag Menu Procesos ADM CEN", IIf(chkOpcion(20).Value, "Verdadero", "Falso"), _
                   "lOpcion21", "Flag Modificar Tipo Tarifa Motorizados", IIf(chkOpcion(21).Value, "Verdadero", "Falso"), "lModulo05", "Flag Modulo Autorizaciones", IIf(chkModulo(5).Value, "Verdadero", "Falso"), "lOpcion22", "Flag Revierte Despacho", IIf(chkOpcion(23).Value, "Verdadero", "Falso"), "lOpcion29", "Flag bloquea emision de notas de credito", IIf(chkOpcion(29).Value, "Verdadero", "Falso"), "lOpcion30", "Anular Documento, despues de haber pasado los días como maximo", IIf(chkOpcion(30).Value, "Verdadero", "Falso"), "lOpcion31", "Modificar Cortesia del Pedido", IIf(chkOpcion(31).Value, "Verdadero", "Falso"), "lOpcion32", "Modificar Canal de Venta del Pedido", IIf(chkOpcion(32).Value, "Verdadero", "Falso"), "lOpcion33", "Activar Autorización de Pedido Sin Mesa", IIf(chkOpcion(33).Value, "Verdadero", "Falso"), "lOpcion34", "Autorizacion de recibo de egreso", IIf(chkOpcion(34).Value, "Verdadero", "Falso"))
                
                    If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
               
                  'Cambiar el SQL
                  Isql = "update TGrupoUsuario set tDetallado ='" & txtDetallado.Text & "', " & _
                         "lModulo01 =" & chkModulo(1).Value & ", lModulo02 =" & chkModulo(2).Value & ", " & _
                         "lModulo03 =" & chkModulo(3).Value & ", lModulo04 =" & chkModulo(4).Value & ", " & _
                         "lOpcion01 =" & chkOpcion(1).Value & ", lOpcion02 =" & chkOpcion(2).Value & ", " & _
                         "lOpcion03 =" & chkOpcion(3).Value & ", lOpcion04 =" & chkOpcion(4).Value & ", " & _
                         "lOpcion05 =" & chkOpcion(5).Value & ", lOpcion06 =" & chkOpcion(6).Value & ", " & _
                         "lOpcion07 =" & chkOpcion(7).Value & ", " & _
                         "lOpcion08 =" & chkOpcion(8).Value & ", " & _
                         "lOpcion09 =" & chkOpcion(9).Value & ", " & _
                         "lOpcion10 =" & chkOpcion(10).Value & ", " & _
                         "lOpcion11 =" & chkOpcion(11).Value & ", " & _
                         "lOpcion12 =" & chkOpcion(12).Value & ", " & _
                         "lOpcion13 =" & chkOpcion(13).Value & ", " & _
                         "lOpcion14 =" & chkOpcion(14).Value & ", " & _
                         "lOpcion15 =" & chkOpcion(15).Value & ", " & _
                         "lOpcion16 =" & chkOpcion(16).Value & ", " & _
                         "lOpcion17 =" & chkOpcion(17).Value & ", " & _
                         "lOpcion18 =" & chkOpcion(18).Value & ", " & _
                         "lOpcion19 =" & chkOpcion(19).Value & ", " & _
                         "lOpcion20 =" & chkOpcion(20).Value & ", " & _
                         "lOpcion21 =" & chkOpcion(21).Value & ",lopcion22=" & chkOpcion(23).Value & ", " & _
                         "lOpcion24=" & chkOpcion(24).Value & ",lOpcion25=" & chkOpcion(25).Value & ", " & _
                         "lopcion26=" & chkOpcion(26).Value & ", lopcion27= " & chkOpcion(27).Value & "," & _
                         "lopcion28= " & chkOpcion(28).Value & ",lopcion29= " & chkOpcion(29).Value & ",lopcion30= " & chkOpcion(30).Value & ",lopcion31= " & chkOpcion(31).Value & ",lopcion32= " & chkOpcion(32).Value & ", lopcion33= " & chkOpcion(33).Value & ", lModulo05 =" & chkModulo(5).Value & ", lopcion34= " & chkOpcion(34).Value & ", " & _
                         "lControlNivel =" & chkOpcion(22).Value & ", lActivo =" & chkActivo.Value & " where tGrupoUsuario = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmGrupoAcceso.RsCabecera.Bookmark
                   frmGrupoAcceso.RsCabecera.Requery
                   If frmGrupoAcceso.RsCabecera.RecordCount = 0 Then
                      frmGrupoAcceso.RsCabecera.Filter = adFilterNone
                   End If
                   frmGrupoAcceso.RsCabecera.Bookmark = nPos
                   GrabaArbol
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmGrupoAcceso.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el GrupoUsuario " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If

               sPasa = txtCodigo.Text
                 
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOUSUARIO", "GRUPO DE USUARIOS", "03", sUsuario, sPasa, "", _
                  "TGRUPOUSUARIO", "Codigo Grupo Usuario", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
               
               If lAuditoria = False Then
                   Screen.MousePointer = vbDefault
                       Exit Sub
               End If

               'Cambia el Delete
               Cn.Execute "delete from TGRUPOUSUARIO where tGrupoUsuario = '" & txtCodigo & "'"
               frmGrupoAcceso.RsCabecera.Requery
               If frmGrupoAcceso.RsCabecera.RecordCount <> 0 Then
                  frmGrupoAcceso.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmGrupoAcceso.RsCabecera.RecordCount = 0, 0, frmGrupoAcceso.RsCabecera.AbsolutePosition) & " de " & frmGrupoAcceso.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
          
   End Select
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Grupo de Usuarios "
    Set RsAcceso = Lib.OpenRecordset("select tCodigoAcceso, tDescripcion, tModulo from TACCESO where tTipoObjeto='MN'", Cn)
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    ssTab.Tab = 0
    cmdTexto.Caption = "Registro " & IIf(frmGrupoAcceso.RsCabecera.RecordCount = 0, 0, frmGrupoAcceso.RsCabecera.AbsolutePosition) & " de " & frmGrupoAcceso.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Public Sub Asignar()
    Dim sCodigo As String
    AgregaArbol
    With frmGrupoAcceso.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tGrupoUsuario), "", !tGrupoUsuario)
        txtDetallado = IIf(IsNull(!tDetallado), "", !tDetallado)
        
        'Check Box
        chkModulo(1) = IIf(!lModulo01 = True, 1, 0) 'Punto de Venta
        chkModulo(2) = IIf(!lModulo02 = True, 1, 0) 'Administracion
        chkModulo(3) = IIf(!lModulo03 = True, 1, 0) 'Reportes
        chkModulo(4) = IIf(!lModulo04 = True, 1, 0) 'auditoria
        chkOpcion(1) = IIf(!lOpcion01 = True, 1, 0) 'Eliminacion Pedido
        chkOpcion(2) = IIf(!lOpcion02 = True, 1, 0) 'Eliminacion Item
        chkOpcion(3) = IIf(!lOpcion03 = True, 1, 0) 'Modifica Precio
        chkOpcion(4) = IIf(!lOpcion04 = True, 1, 0) 'Emision de cortesias
        chkOpcion(5) = IIf(!lOpcion05 = True, 1, 0) 'Anulacion documentos
        chkOpcion(6) = IIf(!lOpcion06 = True, 1, 0) 'Anulacion doc. otro turno
        chkOpcion(7) = IIf(!lOpcion07 = True, 1, 0) 'Anulacion Recibos
        chkOpcion(8) = IIf(!lOpcion08 = True, 1, 0) 'Cuenta por Cobrar
        chkOpcion(9) = IIf(!lOpcion09 = True, 1, 0) 'Cuenta Corriente
        chkOpcion(10) = IIf(!lOpcion10 = True, 1, 0) 'Desc/Recar Documentos
        chkOpcion(11) = IIf(!lOpcion11 = True, 1, 0) 'Obliga Cierre
        chkOpcion(12) = IIf(!lOpcion12 = True, 1, 0) 'Reimpresiones
        chkOpcion(13) = IIf(!lOpcion13 = True, 1, 0) 'Anulación Cta.Cte.
        chkOpcion(14) = IIf(!lOpcion14 = True, 1, 0) 'Impresion de Precuentas
        chkOpcion(15) = IIf(!lOpcion15 = True, 1, 0) 'Transferencias
        chkOpcion(16) = IIf(!lOpcion16 = True, 1, 0) 'MultiCajero
        chkOpcion(17) = IIf(!lOpcion17 = True, 1, 0) 'administrador en turnos
        chkOpcion(18) = IIf(!lOpcion18 = True, 1, 0) 'controler en turnos
        chkOpcion(19) = IIf(!lOpcion19 = True, 1, 0) 'administracion centralizada configuraciones
        chkOpcion(20) = IIf(!lOpcion20 = True, 1, 0) 'administracion centralizada ejeucion de procesos
        chkOpcion(21) = IIf(!LOPCION21 = True, 1, 0) 'PLANILLA MOTORIZADOS
        chkModulo(5) = IIf(!lModulo05 = True, 1, 0) 'autorizaciones
        chkOpcion(22) = IIf(!lControlNivel = True, 1, 0) 'Control de Nivel de Usuario
        chkOpcion(23) = IIf(!lopcion22 = True, 1, 0)
        chkOpcion(24) = IIf(!lOpcion24 = True, 1, 0) 'Autoriza envios de productos a produccion
        chkOpcion(25) = IIf(!lOpcion25 = True, 1, 0)
        chkOpcion(26) = IIf(!lOpcion26 = True, 1, 0) ' modificacion de numeros de pax y mesa
        chkOpcion(27) = IIf(!lOpcion27 = True, 1, 0) ' emitir notas de credito
        chkOpcion(28) = IIf(!lOpcion28 = True, 1, 0) ' Activar Clave Permiso para Pago Otros
        chkOpcion(29) = IIf(!lOpcion29 = True, 1, 0) ' Bloquea notas de credito en documentos fuera del mes
        chkOpcion(30) = IIf(!lOpcion30 = True, 1, 0) ' Anular Documento, despues de haber pasado los días como maximo
        
        chkOpcion(31) = IIf(!lOpcion31 = True, 1, 0) ' Modificar cortesia del pedido
        chkOpcion(32) = IIf(IsNull(!lOpcion32), 0, IIf(!lOpcion32 = True, 1, 0)) ' Modificar Canal de venta
        chkOpcion(33) = IIf(IsNull(!lOpcion33), 0, IIf(!lOpcion33 = True, 1, 0)) ' Activar Autorización de Pedido Sin Mesa
        chkOpcion(34) = IIf(IsNull(!lOpcion34), 0, IIf(!lOpcion34 = True, 1, 0))
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    QuitaArbol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrupoAccesoDetalle = Nothing
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
End Sub

Public Sub Iniciar()
   chkModulo(1).Value = 0
   chkModulo(2).Value = 0
   chkModulo(3).Value = 0
   chkModulo(4).Value = 0
   chkOpcion(1).Value = 0
   chkOpcion(2).Value = 0
   chkOpcion(3).Value = 0
   chkOpcion(4).Value = 0
   chkOpcion(5).Value = 0
   chkOpcion(6).Value = 0
   chkOpcion(7).Value = 0
   chkOpcion(8).Value = 0
   chkOpcion(9).Value = 0
   chkOpcion(10).Value = 0
   chkOpcion(11).Value = 0
   chkOpcion(12).Value = 0
   chkOpcion(13).Value = 0
   chkOpcion(14).Value = 0
   chkOpcion(15).Value = 0
   chkOpcion(16).Value = 0
   chkOpcion(17).Value = 0
   chkOpcion(18).Value = 0
   chkActivo.Value = 1
   chkOpcion(19).Value = 0
   chkOpcion(20).Value = 0
   chkOpcion(21).Value = 0
   chkOpcion(23).Value = 0
End Sub

Private Sub tvAdministracion_NodeCheck(ByVal Node As MSComctlLib.Node)
   Node.Checked = True
End Sub

Public Sub AgregaArbol()
    
    Dim xNode As Node
    Dim yNode As Node
    
    RsAcceso.Filter = "tModulo = '02'"
    tvPuntoVenta.Nodes.Clear
    
    Do While Not RsAcceso.EOF
       If Mid(RsAcceso!tCodigoAcceso, 4, 5) = "00000" Then
          Set xNode = tvPuntoVenta.Nodes.Add(, , "M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), RsAcceso!tDescripcion)
          xNode.Checked = True
          xNode.Expanded = True
       ElseIf Mid(RsAcceso!tCodigoAcceso, 7, 2) = "00" Then
          Set xNode = tvPuntoVenta.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), RsAcceso!tDescripcion)
          xNode.Checked = True
          xNode.Expanded = True
       Else
          Set xNode = tvPuntoVenta.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 7), RsAcceso!tDescripcion)
          xNode.Checked = True
       End If
       RsAcceso.MoveNext
    Loop
    
    RsAcceso.Filter = "tModulo = '03'"
    tvAdministracion.Nodes.Clear
    Do While Not RsAcceso.EOF
       If Mid(RsAcceso!tCodigoAcceso, 4, 5) = "00000" Then
          Set xNode = tvAdministracion.Nodes.Add(, , "M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), RsAcceso!tDescripcion)
          xNode.Checked = True
          xNode.Expanded = True
       ElseIf Mid(RsAcceso!tCodigoAcceso, 7, 2) = "00" Then
          Set xNode = tvAdministracion.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), RsAcceso!tDescripcion)
          xNode.Checked = True
          xNode.Expanded = True
       Else
          Set xNode = tvAdministracion.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 7), RsAcceso!tDescripcion)
          xNode.Checked = True
       End If
       RsAcceso.MoveNext
    Loop
    
    RsAcceso.Filter = "tModulo = '04'"
    tvConsulta.Nodes.Clear
    Do While Not RsAcceso.EOF
       If Mid(RsAcceso!tCodigoAcceso, 4, 5) = "00000" Then
          Set yNode = tvConsulta.Nodes.Add(, , "M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), RsAcceso!tDescripcion)
          yNode.Checked = True
          yNode.Expanded = True
       ElseIf Mid(RsAcceso!tCodigoAcceso, 7, 2) = "00" Then
          Set yNode = tvConsulta.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), RsAcceso!tDescripcion)
          yNode.Checked = True
          yNode.Expanded = True
       Else
          Set yNode = tvConsulta.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 7), RsAcceso!tDescripcion)
          yNode.Checked = True
       End If
       RsAcceso.MoveNext
    Loop
    chkBotonP1.Value = 1
    chkBotonP2.Value = 1
    chkBotonP3.Value = 1
    chkBotonP4.Value = 1
    chkBotonP5.Value = 1
    chkBotonP6.Value = 1
    chkBotonP7.Value = 1
    chkBotonP8.Value = 1
    chkBotonP9.Value = 1
    chkBotonP10.Value = 1
    chkBotonP11.Value = 1
    chkBotonP12.Value = 1
    chkBotonP13.Value = 1
    chkBotonP14.Value = 1
    chkBotonA1.Value = 1
    chkBotonA2.Value = 1
    chkBotonA3.Value = 1
    chkBotonA4.Value = 1
    chkBotonA5.Value = 1
    chkBotonA6.Value = 1
    chkBotonC1.Value = 1
    chkBotonC2.Value = 1
    chkBotonC3.Value = 1
    chkBotonC4.Value = 1
    chkBotonC5.Value = 1
    chkBotonC6.Value = 1
    chkBotonC7.Value = 1
End Sub

Public Sub GrabaArbol()
    RsAcceso.Filter = "tModulo = '02'"
    Dim i As Integer
    Dim sCodigo As String
    
        'Inserta Movimiento auditoria
'    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "03", sUsuario, txtCodigo.Text, "", _
'       "TGRUPOUSUARIO", "Codigo Grupo Usuario", txtCodigo.Text)
'    If lAuditoria = False Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If

    Cn.Execute "delete from TGRUPOACCESO where tGrupoUsuario='" & txtCodigo & "'"
    
    RsAcceso.Filter = "tModulo = '02'"
    For i = 1 To tvPuntoVenta.Nodes.Count
        If Not tvPuntoVenta.Nodes.Item(i).Checked Then
            sCodigo = Left("3" & Mid(tvPuntoVenta.Nodes(i).Key, 2) + "0000000", 8)
            
            lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, sCodigo, _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", sCodigo)
            
            Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '" & sCodigo & "')"
        End If
    Next i
    RsAcceso.Filter = "tModulo = '03'"
    For i = 1 To tvAdministracion.Nodes.Count
        If Not tvAdministracion.Nodes.Item(i).Checked Then
            sCodigo = Left("1" & Mid(tvAdministracion.Nodes(i).Key, 2) + "0000000", 8)
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, sCodigo, _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", sCodigo)
            Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '" & sCodigo & "')"
        End If
    Next i
    RsAcceso.Filter = "tModulo = '04'"
    For i = 1 To tvConsulta.Nodes.Count
        If Not tvConsulta.Nodes.Item(i).Checked Then
            sCodigo = Left("2" & Mid(tvConsulta.Nodes(i).Key, 2) + "0000000", 8)
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, sCodigo, _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", sCodigo)
            Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '" & sCodigo & "')"
        End If
    Next i
    
    If chkBotonA1.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000001", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000001")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000001')"
    End If
    If chkBotonA2.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000002", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000002")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000002')"
    End If
    If chkBotonA3.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000003", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000003")
    
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000003')"
    End If
    If chkBotonA4.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000004", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000004")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000004')"
    End If
    If chkBotonA5.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000005", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000005")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000005')"
    End If
    If chkBotonA6.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000006", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000006")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000006')"
    End If
    If chkBotonC1.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000007", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000007")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000007')"
    End If
    If chkBotonC2.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000008", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000008")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000008')"
    End If
    If chkBotonC3.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000009", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000009")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000009')"
    End If
    If chkBotonC4.Value = 0 Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000010", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000010")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000010')"
    End If
    If chkBotonC5.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000011", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000011")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000011')"
    End If
    If chkBotonC6.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000012", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000012")

       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000012')"
    End If
    If chkBotonC7.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000013", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000013")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000013')"
    End If
    If chkBotonP1.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000014", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000014")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000014')"
    End If
    If chkBotonP2.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000015", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000015")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000015')"
    End If
    If chkBotonP3.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000016", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000016")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000016')"
    End If
    If chkBotonP4.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000017", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000017")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000017')"
    End If
    If chkBotonP5.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000018", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000018")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000018')"
    End If
    If chkBotonP6.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000019", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000019")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000019')"
    End If
    If chkBotonP7.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000020", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000020")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000020')"
    End If
    If chkBotonP8.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000021", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000021")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000021')"
    End If
    If chkBotonP9.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000022", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000022")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000022')"
    End If
    If chkBotonP10.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000023", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000023")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000023')"
    End If
    If chkBotonP11.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000024", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000024")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000024')"
    End If
    If chkBotonP12.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000025", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000025")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000025')"
    End If
    If chkBotonP13.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000026", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000026")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000026')"
    End If
    If chkBotonP14.Value = 0 Then
    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TGRUPOACCESO", "GRUPO ACCESO", "01", sUsuario, txtCodigo.Text, "00000027", _
       "TGRUPOUSUARIO", "Codigo Grupo", txtCodigo.Text, "tcodigoacceso", "Codigo Acceso", "00000027")
       Cn.Execute "Insert into TGRUPOACCESO values ('" & txtCodigo & "', '00000027')"
    End If
    
End Sub

Public Sub QuitaArbol()
    Isql = "SELECT dbo.TGRUPOACCESO.tCodigoAcceso, dbo.TACCESO.tModulo FROM dbo.TACCESO INNER JOIN dbo.TGRUPOACCESO ON dbo.TACCESO.tCodigoAcceso = dbo.TGRUPOACCESO.tCodigoAcceso WHERE (dbo.TACCESO.tTipoObjeto = N'MN') AND (dbo.TGRUPOACCESO.tGrupoUsuario = N'" & txtCodigo & "')"
    Set RsGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
    Do While Not RsGrupoAcceso.EOF
       If Mid(RsGrupoAcceso!tCodigoAcceso, 4, 5) = "00000" Then
          sCodigo = "M" & Mid(RsGrupoAcceso!tCodigoAcceso, 2, 2)
       ElseIf Mid(RsGrupoAcceso!tCodigoAcceso, 7, 2) = "00" Then
          sCodigo = "M" & Mid(RsGrupoAcceso!tCodigoAcceso, 2, 5)
       Else
          sCodigo = "M" & Mid(RsGrupoAcceso!tCodigoAcceso, 2, 7)
       End If
       
       If RsGrupoAcceso!tModulo = "03" Then
          tvAdministracion.Nodes.Item(sCodigo).Checked = False
       ElseIf RsGrupoAcceso!tModulo = "04" Then
          tvConsulta.Nodes.Item(sCodigo).Checked = False
       Else
          tvPuntoVenta.Nodes.Item(sCodigo).Checked = False
       End If
       RsGrupoAcceso.MoveNext
    Loop
    
    Isql = "SELECT dbo.TGRUPOACCESO.tCodigoAcceso, dbo.TACCESO.tModulo FROM dbo.TACCESO INNER JOIN dbo.TGRUPOACCESO ON dbo.TACCESO.tCodigoAcceso = dbo.TGRUPOACCESO.tCodigoAcceso WHERE (dbo.TACCESO.tTipoObjeto = N'BT') AND (dbo.TGRUPOACCESO.tGrupoUsuario = N'" & txtCodigo & "')"
    Set RsGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
    Do While Not RsGrupoAcceso.EOF
       Select Case RsGrupoAcceso!tCodigoAcceso
              Case Is = "00000001"
                   chkBotonA1.Value = 0
              Case Is = "00000002"
                   chkBotonA2.Value = 0
              Case Is = "00000003"
                   chkBotonA3.Value = 0
              Case Is = "00000004"
                   chkBotonA4.Value = 0
              Case Is = "00000005"
                   chkBotonA5.Value = 0
              Case Is = "00000006"
                   chkBotonA6.Value = 0
              Case Is = "00000007"
                   chkBotonC1.Value = 0
              Case Is = "00000008"
                   chkBotonC2.Value = 0
              Case Is = "00000009"
                   chkBotonC3.Value = 0
              Case Is = "00000010"
                   chkBotonC4.Value = 0
              Case Is = "00000011"
                   chkBotonC5.Value = 0
              Case Is = "00000012"
                   chkBotonC6.Value = 0
              Case Is = "00000013"
                   chkBotonC7.Value = 0
              Case Is = "00000014"
                   chkBotonP1.Value = 0
              Case Is = "00000015"
                   chkBotonP2.Value = 0
              Case Is = "00000016"
                   chkBotonP3.Value = 0
              Case Is = "00000017"
                   chkBotonP4.Value = 0
              Case Is = "00000018"
                   chkBotonP5.Value = 0
              Case Is = "00000019"
                   chkBotonP6.Value = 0
              Case Is = "00000020"
                   chkBotonP7.Value = 0
              Case Is = "00000021"
                   chkBotonP8.Value = 0
              Case Is = "00000022"
                   chkBotonP9.Value = 0
              Case Is = "00000023"
                   chkBotonP10.Value = 0
              Case Is = "00000024"
                   chkBotonP11.Value = 0
              Case Is = "00000025"
                   chkBotonP12.Value = 0
              Case Is = "00000026"
                   chkBotonP13.Value = 0
              Case Is = "00000027"
                   chkBotonP14.Value = 0
       End Select
       RsGrupoAcceso.MoveNext
    Loop
    
End Sub



