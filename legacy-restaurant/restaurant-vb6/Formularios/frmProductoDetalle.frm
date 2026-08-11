VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   2835
   ClientTop       =   6105
   ClientWidth     =   11835
   Icon            =   "frmProductoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBoton 
      Caption         =   " Botonera "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7410
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2325
      Begin VB.CommandButton cmdBoton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1950
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1950
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1950
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2760
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2760
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2760
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3570
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3570
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3570
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4380
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4380
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4380
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5190
         Width           =   720
      End
      Begin VB.CommandButton cmdNA 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1530
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5190
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5190
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         BackColor       =   &H80000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   6330
         Width           =   510
      End
      Begin VB.Label txtBoton 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   870
         TabIndex        =   35
         Top             =   6300
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11775
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7425
      Width           =   11835
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Caja Rápida"
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
         Index           =   4
         Left            =   9324
         Picture         =   "frmProductoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5550
         TabIndex        =   6
         Top             =   45
         Width           =   5610
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4110
            Picture         =   "frmProductoDetalle.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4590
            Picture         =   "frmProductoDetalle.frx":0ACE
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5070
            Picture         =   "frmProductoDetalle.frx":1010
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmProductoDetalle.frx":1552
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmProductoDetalle.frx":1A94
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmProductoDetalle.frx":1FD6
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
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
            Left            =   1485
            TabIndex        =   37
            Top             =   180
            Width           =   2520
         End
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
         Left            =   6910
         Picture         =   "frmProductoDetalle.frx":2518
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
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
         Left            =   5715
         Picture         =   "frmProductoDetalle.frx":2A4A
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   45
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
         Left            =   8121
         Picture         =   "frmProductoDetalle.frx":2F7C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   1170
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
         Index           =   3
         Left            =   10560
         Picture         =   "frmProductoDetalle.frx":307E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab tabProducto 
      Height          =   3435
      Left            =   2340
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3960
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   6059
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabMaxWidth     =   18
      TabCaption(0)   =   "Precios"
      TabPicture(0)   =   "frmProductoDetalle.frx":3170
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkLlevar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDelivery"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkVisualizarLocal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkVisualizarDelivery"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkVisualizarLlevar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Precios (2)"
      TabPicture(1)   =   "frmProductoDetalle.frx":318C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "chkCanal4"
      Tab(1).Control(3)=   "chkCanal5"
      Tab(1).Control(4)=   "chkvisualizarCanal4"
      Tab(1).Control(5)=   "chkVisualizarCanal5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Costos"
      TabPicture(2)   =   "frmProductoDetalle.frx":31A8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label(8)"
      Tab(2).Control(1)=   "Label(5)"
      Tab(2).Control(2)=   "Label(14)"
      Tab(2).Control(3)=   "lblDirecto"
      Tab(2).Control(4)=   "lblReceta"
      Tab(2).Control(5)=   "TabCostos"
      Tab(2).Control(6)=   "cboDescargo"
      Tab(2).Control(7)=   "cmdBusca(0)"
      Tab(2).Control(8)=   "txtReceta"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtDirecto"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdBusca(1)"
      Tab(2).Control(11)=   "txtEnlace"
      Tab(2).Control(12)=   "cmdOpcion(5)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Areas Impresión"
      TabPicture(3)   =   "frmProductoDetalle.frx":31C4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label(20)"
      Tab(3).Control(1)=   "cboImpresion"
      Tab(3).Control(2)=   "grdGrilla"
      Tab(3).Control(3)=   "cmdArea(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdArea(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Ofertas"
      TabPicture(4)   =   "frmProductoDetalle.frx":31E0
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkAcumulable"
      Tab(4).Control(1)=   "txtNombre"
      Tab(4).Control(2)=   "ChkActv"
      Tab(4).Control(3)=   "txtRatio"
      Tab(4).Control(4)=   "cmdOferta"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "grdOferta"
      Tab(4).Control(6)=   "dtpHoraFin"
      Tab(4).Control(7)=   "dtpFecha"
      Tab(4).Control(8)=   "dtpHoraIni"
      Tab(4).Control(9)=   "Label(25)"
      Tab(4).Control(10)=   "Label10"
      Tab(4).Control(11)=   "Label(24)"
      Tab(4).Control(12)=   "Label(23)"
      Tab(4).Control(13)=   "Label(22)"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "Insumo Critico"
      TabPicture(5)   =   "frmProductoDetalle.frx":31FC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkInsumoCritico"
      Tab(5).Control(1)=   "fraInsumo"
      Tab(5).ControlCount=   2
      Begin VB.CheckBox chkInsumoCritico 
         Alignment       =   1  'Right Justify
         Caption         =   "Control Insumo Crítico  :"
         Height          =   225
         Left            =   -74820
         TabIndex        =   221
         Top             =   540
         Width           =   2565
      End
      Begin VB.Frame fraInsumo 
         Caption         =   "Insumo"
         Enabled         =   0   'False
         Height          =   1935
         Left            =   -74820
         TabIndex        =   217
         Top             =   900
         Width           =   6975
         Begin VB.CommandButton cmdBuscaInsumo 
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
            Left            =   4920
            Picture         =   "frmProductoDetalle.frx":3218
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   720
            Width           =   705
         End
         Begin VB.Label lblInsumo 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   1530
            TabIndex        =   220
            Top             =   720
            Width           =   3360
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Index           =   60
            Left            =   480
            TabIndex        =   219
            Top             =   780
            Width           =   930
         End
      End
      Begin VB.CheckBox chkAcumulable 
         Alignment       =   1  'Right Justify
         Caption         =   "Acumulable :"
         Height          =   195
         Left            =   -74745
         TabIndex        =   207
         Top             =   2445
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   -73710
         MaxLength       =   50
         TabIndex        =   206
         Text            =   " "
         Top             =   450
         Width           =   6135
      End
      Begin VB.CheckBox ChkActv 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   -74385
         TabIndex        =   205
         Top             =   2715
         Width           =   855
      End
      Begin VB.TextBox txtRatio 
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
         Left            =   -73710
         MaxLength       =   15
         TabIndex        =   204
         Text            =   "0.00"
         Top             =   2085
         Width           =   1740
      End
      Begin VB.CommandButton cmdOferta 
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
         Left            =   -73170
         Picture         =   "frmProductoDetalle.frx":331A
         Style           =   1  'Graphical
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   2715
         Width           =   1170
      End
      Begin VB.CheckBox chkVisualizarCanal5 
         Caption         =   "Visualizar"
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
         Left            =   -70995
         TabIndex        =   136
         Top             =   3150
         Width           =   2850
      End
      Begin VB.CheckBox chkvisualizarCanal4 
         Caption         =   "Visualizar"
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
         Left            =   -74685
         TabIndex        =   135
         Top             =   3150
         Width           =   2850
      End
      Begin VB.CheckBox chkCanal5 
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
         Left            =   -70725
         TabIndex        =   134
         Top             =   450
         Width           =   2940
      End
      Begin VB.CheckBox chkCanal4 
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
         Left            =   -74460
         TabIndex        =   133
         Top             =   450
         Width           =   2940
      End
      Begin VB.Frame Frame5 
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -70995
         TabIndex        =   117
         Top             =   450
         Width           =   3435
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
            Left            =   1395
            MaxLength       =   15
            TabIndex        =   121
            Text            =   " "
            Top             =   2205
            Width           =   1275
         End
         Begin VB.CheckBox chkImpuesto2Canal5 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1072
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto3Canal5 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   1417
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto1Canal5 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   727
            Width           =   210
         End
         Begin MSDataListLib.DataCombo cboMonedaCanal5 
            Height          =   315
            Left            =   1395
            TabIndex        =   122
            Top             =   1830
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label txtImpuesto1Canal5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   132
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label txtImpuesto3Canal5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   131
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label txtImpuesto2Canal5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   130
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label txtPrecioBaseCanal5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   129
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lblI2Canal5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   360
            TabIndex        =   128
            Top             =   1095
            Width           =   870
         End
         Begin VB.Label lblI3Canal5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto3 :"
            Height          =   195
            Left            =   405
            TabIndex        =   127
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblI1Canal5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   360
            TabIndex        =   126
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            Height          =   195
            Index           =   38
            Left            =   300
            TabIndex        =   125
            Top             =   405
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Index           =   37
            Left            =   225
            TabIndex        =   124
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Index           =   36
            Left            =   555
            TabIndex        =   123
            Top             =   1890
            Width           =   675
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -74685
         TabIndex        =   101
         Top             =   450
         Width           =   3435
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
            Left            =   1395
            MaxLength       =   15
            TabIndex        =   105
            Text            =   " "
            Top             =   2205
            Width           =   1275
         End
         Begin VB.CheckBox chkImpuesto2Canal4 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1072
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto3Canal4 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1417
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto1Canal4 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2775
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   727
            Width           =   210
         End
         Begin MSDataListLib.DataCombo cboMonedaCanal4 
            Height          =   315
            Left            =   1395
            TabIndex        =   106
            Top             =   1830
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label txtImpuesto1Canal4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   116
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label txtImpuesto3Canal4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   115
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label txtImpuesto2Canal4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   114
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label txtPrecioBaseCanal4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1395
            TabIndex        =   113
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lblI2Canal4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   360
            TabIndex        =   112
            Top             =   1095
            Width           =   870
         End
         Begin VB.Label lblI3Canal4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto3 :"
            Height          =   195
            Left            =   405
            TabIndex        =   111
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblI1Canal4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   360
            TabIndex        =   110
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            Height          =   195
            Index           =   35
            Left            =   300
            TabIndex        =   109
            Top             =   405
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Index           =   34
            Left            =   225
            TabIndex        =   108
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Index           =   33
            Left            =   555
            TabIndex        =   107
            Top             =   1890
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ver Receta"
         Height          =   435
         Index           =   5
         Left            =   -72615
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1155
      End
      Begin VB.TextBox txtEnlace 
         BackColor       =   &H80000016&
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
         Height          =   315
         Left            =   -74865
         MaxLength       =   7
         TabIndex        =   99
         Text            =   " "
         Top             =   2925
         Visible         =   0   'False
         Width           =   2130
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
         Height          =   315
         Index           =   1
         Left            =   -72165
         Picture         =   "frmProductoDetalle.frx":341C
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1995
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtDirecto 
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
         Height          =   315
         Left            =   -73470
         MaxLength       =   24
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1050
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReceta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -73470
         MaxLength       =   24
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1995
         Visible         =   0   'False
         Width           =   1215
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
         Height          =   315
         Index           =   0
         Left            =   -72165
         Picture         =   "frmProductoDetalle.frx":351E
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1050
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdArea 
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
         Index           =   1
         Left            =   -74445
         Picture         =   "frmProductoDetalle.frx":3620
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   990
         Width           =   1170
      End
      Begin VB.CommandButton cmdArea 
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
         Left            =   -74445
         Picture         =   "frmProductoDetalle.frx":3B52
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1170
      End
      Begin VB.CheckBox chkVisualizarLlevar 
         Caption         =   "Visualizar"
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
         Left            =   6300
         TabIndex        =   92
         Top             =   3135
         Width           =   2850
      End
      Begin VB.CheckBox chkVisualizarDelivery 
         Caption         =   "Visualizar"
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
         Left            =   3195
         TabIndex        =   91
         Top             =   3135
         Width           =   2850
      End
      Begin VB.CheckBox chkVisualizarLocal 
         Caption         =   "Visualizar"
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
         Left            =   135
         TabIndex        =   90
         Top             =   3135
         Width           =   2850
      End
      Begin VB.CheckBox chkDelivery 
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
         Left            =   3375
         TabIndex        =   89
         Top             =   480
         Width           =   2670
      End
      Begin VB.CheckBox chkLlevar 
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
         Left            =   6525
         TabIndex        =   88
         Top             =   480
         Width           =   2625
      End
      Begin VB.Frame Frame2 
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   3195
         TabIndex        =   72
         Top             =   435
         Width           =   3030
         Begin VB.CheckBox chkImpuesto1Delivery 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   727
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto3Delivery 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1417
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto2Delivery 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1072
            Width           =   210
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
            Left            =   1305
            MaxLength       =   15
            TabIndex        =   73
            Text            =   " "
            Top             =   2205
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo cboMonedaDelivery 
            Height          =   315
            Left            =   1305
            TabIndex        =   77
            Top             =   1830
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Index           =   29
            Left            =   510
            TabIndex        =   87
            Top             =   1890
            Width           =   675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Index           =   27
            Left            =   180
            TabIndex        =   86
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            Height          =   195
            Index           =   26
            Left            =   255
            TabIndex        =   85
            Top             =   405
            Width           =   930
         End
         Begin VB.Label lblI1Delivery 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   315
            TabIndex        =   84
            Top             =   750
            Width           =   870
         End
         Begin VB.Label lblI3Delivery 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto3 :"
            Height          =   195
            Left            =   360
            TabIndex        =   83
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblI2Delivery 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   315
            TabIndex        =   82
            Top             =   1095
            Width           =   870
         End
         Begin VB.Label txtPrecioBaseDelivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   81
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label txtImpuesto2Delivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   80
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label txtImpuesto3Delivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   79
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label txtImpuesto1Delivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   78
            Top             =   705
            Width           =   1275
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   6300
         TabIndex        =   56
         Top             =   435
         Width           =   3030
         Begin VB.CheckBox chkImpuesto1Llevar 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   727
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto3Llevar 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1417
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto2Llevar 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1072
            Width           =   210
         End
         Begin VB.TextBox txtLlevar 
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
            Left            =   1305
            MaxLength       =   15
            TabIndex        =   57
            Text            =   " "
            Top             =   2205
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo cboMonedaLlevar 
            Height          =   315
            Left            =   1305
            TabIndex        =   61
            Top             =   1830
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Index           =   30
            Left            =   510
            TabIndex        =   71
            Top             =   1890
            Width           =   675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   70
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            Height          =   195
            Index           =   10
            Left            =   255
            TabIndex        =   69
            Top             =   405
            Width           =   930
         End
         Begin VB.Label lblI1Llevar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   315
            TabIndex        =   68
            Top             =   750
            Width           =   870
         End
         Begin VB.Label lblI3Llevar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto3 :"
            Height          =   195
            Left            =   360
            TabIndex        =   67
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblI2Llevar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   315
            TabIndex        =   66
            Top             =   1095
            Width           =   870
         End
         Begin VB.Label txtPrecioBaseLlevar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   65
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label txtImpuesto2Llevar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   64
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label txtImpuesto3Llevar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   63
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label txtImpuesto1Llevar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   62
            Top             =   705
            Width           =   1275
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   2160
         Left            =   -73050
         TabIndex        =   137
         Top             =   990
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3810
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
      Begin MSDataListLib.DataCombo cboImpresion 
         Height          =   315
         Left            =   -73050
         TabIndex        =   138
         Top             =   540
         Width           =   2505
         _ExtentX        =   4419
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
      Begin MSDataListLib.DataCombo cboDescargo 
         Height          =   315
         Left            =   -73470
         TabIndex        =   139
         Top             =   495
         Width           =   1995
         _ExtentX        =   3519
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
      Begin TabDlg.SSTab TabCostos 
         Height          =   2985
         Left            =   -71355
         TabIndex        =   140
         Top             =   405
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   5265
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "n/d"
         TabPicture(0)   =   "frmProductoDetalle.frx":3C54
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(18)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label(17)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(16)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(15)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(9)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "mscGrafico"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtUtilidad"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtManoObra"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtGasto"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtInsumo"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtPrecioCosto"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "n/d"
         TabPicture(1)   =   "frmProductoDetalle.frx":3C70
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label(43)"
         Tab(1).Control(1)=   "Label(42)"
         Tab(1).Control(2)=   "Label(41)"
         Tab(1).Control(3)=   "Label(40)"
         Tab(1).Control(4)=   "Label(39)"
         Tab(1).Control(5)=   "mscGrafico2"
         Tab(1).Control(6)=   "txtPrecioCosto2"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtInsumo2"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtGasto2"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtManoObra2"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txtUtilidad2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).ControlCount=   11
         TabCaption(2)   =   "n/d"
         TabPicture(2)   =   "frmProductoDetalle.frx":3C8C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label(48)"
         Tab(2).Control(1)=   "Label(47)"
         Tab(2).Control(2)=   "Label(46)"
         Tab(2).Control(3)=   "Label(45)"
         Tab(2).Control(4)=   "Label(44)"
         Tab(2).Control(5)=   "mscGrafico3"
         Tab(2).Control(6)=   "txtPrecioCosto3"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "txtInsumo3"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "txtGasto3"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "txtManoObra3"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "txtUtilidad3"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).ControlCount=   11
         TabCaption(3)   =   "n/d"
         TabPicture(3)   =   "frmProductoDetalle.frx":3CA8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label(53)"
         Tab(3).Control(1)=   "Label(52)"
         Tab(3).Control(2)=   "Label(51)"
         Tab(3).Control(3)=   "Label(50)"
         Tab(3).Control(4)=   "Label(49)"
         Tab(3).Control(5)=   "mscGrafico4"
         Tab(3).Control(6)=   "txtPrecioCosto4"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "txtInsumo4"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "txtGasto4"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "txtManoObra4"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "txtUtilidad4"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).ControlCount=   11
         TabCaption(4)   =   "n/d"
         TabPicture(4)   =   "frmProductoDetalle.frx":3CC4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtUtilidad5"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "txtManoObra5"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "txtGasto5"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "txtInsumo5"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "txtPrecioCosto5"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "mscGrafico5"
         Tab(4).Control(6)=   "Label(54)"
         Tab(4).Control(7)=   "Label(55)"
         Tab(4).Control(8)=   "Label(56)"
         Tab(4).Control(9)=   "Label(57)"
         Tab(4).Control(10)=   "Label(58)"
         Tab(4).ControlCount=   11
         Begin VB.TextBox txtPrecioCosto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   165
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtInsumo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   164
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   163
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtManoObra 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   162
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   161
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   160
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtManoObra2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   159
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtGasto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   158
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtInsumo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   157
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   156
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   155
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtManoObra3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   154
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtGasto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   153
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtInsumo3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   152
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   151
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   150
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtManoObra4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   149
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtGasto4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   148
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtInsumo4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   147
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   146
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   145
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtManoObra5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   144
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtGasto5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   143
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtInsumo5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   142
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   141
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin MSChart20Lib.MSChart mscGrafico2 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmProductoDetalle.frx":3CE0
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico3 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmProductoDetalle.frx":6581
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico4 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmProductoDetalle.frx":8E06
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico5 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmProductoDetalle.frx":B68B
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico 
            Height          =   2550
            Left            =   2415
            OleObjectBlob   =   "frmProductoDetalle.frx":DF10
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   9
            Left            =   135
            TabIndex        =   195
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   15
            Left            =   525
            TabIndex        =   194
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   193
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Height          =   195
            Index           =   17
            Left            =   285
            TabIndex        =   192
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   18
            Left            =   570
            TabIndex        =   191
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   39
            Left            =   -74430
            TabIndex        =   190
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Height          =   195
            Index           =   40
            Left            =   -74715
            TabIndex        =   189
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   41
            Left            =   -74865
            TabIndex        =   188
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   42
            Left            =   -74475
            TabIndex        =   187
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   43
            Left            =   -74865
            TabIndex        =   186
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   44
            Left            =   -74430
            TabIndex        =   185
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Height          =   195
            Index           =   45
            Left            =   -74715
            TabIndex        =   184
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   46
            Left            =   -74865
            TabIndex        =   183
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   47
            Left            =   -74475
            TabIndex        =   182
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   48
            Left            =   -74865
            TabIndex        =   181
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   49
            Left            =   -74430
            TabIndex        =   180
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Height          =   195
            Index           =   50
            Left            =   -74715
            TabIndex        =   179
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   51
            Left            =   -74865
            TabIndex        =   178
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   52
            Left            =   -74475
            TabIndex        =   177
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   53
            Left            =   -74865
            TabIndex        =   176
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   54
            Left            =   -74430
            TabIndex        =   175
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Height          =   195
            Index           =   55
            Left            =   -74715
            TabIndex        =   174
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   56
            Left            =   -74865
            TabIndex        =   173
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   57
            Left            =   -74475
            TabIndex        =   172
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   58
            Left            =   -74865
            TabIndex        =   171
            Top             =   2070
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "        "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   90
         TabIndex        =   40
         Top             =   450
         Width           =   3030
         Begin VB.TextBox txtPrecioVenta 
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
            Left            =   1305
            MaxLength       =   15
            TabIndex        =   44
            Text            =   " "
            Top             =   2205
            Width           =   1275
         End
         Begin VB.CheckBox chkImpuesto2 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1072
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto3 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1417
            Width           =   210
         End
         Begin VB.CheckBox chkImpuesto1 
            Caption         =   "Check1"
            Height          =   240
            Left            =   2685
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   727
            Width           =   210
         End
         Begin MSDataListLib.DataCombo cboMoneda 
            Height          =   315
            Left            =   1305
            TabIndex        =   45
            Top             =   1830
            Width           =   795
            _ExtentX        =   1402
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
         Begin VB.Label lblLocal 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   225
            TabIndex        =   202
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Index           =   28
            Left            =   510
            TabIndex        =   55
            Top             =   1890
            Width           =   675
         End
         Begin VB.Label txtImpuesto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   54
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label txtImpuesto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   53
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label txtImpuesto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   52
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label txtPrecioBase 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1305
            TabIndex        =   51
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lblI2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   315
            TabIndex        =   50
            Top             =   1095
            Width           =   870
         End
         Begin VB.Label lblI3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto3 :"
            Height          =   195
            Left            =   360
            TabIndex        =   49
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblI1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   315
            TabIndex        =   48
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            Height          =   195
            Index           =   13
            Left            =   255
            TabIndex        =   47
            Top             =   405
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Index           =   12
            Left            =   180
            TabIndex        =   46
            Top             =   2250
            Width           =   1005
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdOferta 
         Height          =   2145
         Left            =   -71910
         TabIndex        =   208
         Top             =   1185
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   3784
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
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   375
         Left            =   -73710
         TabIndex        =   209
         Top             =   1635
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
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
         Format          =   55246851
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   -73725
         TabIndex        =   210
         Top             =   810
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM"
         Format          =   55246851
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   375
         Left            =   -73680
         TabIndex        =   211
         Top             =   1200
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
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
         Format          =   55246851
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Ratio :"
         Height          =   195
         Index           =   25
         Left            =   -74250
         TabIndex        =   216
         Top             =   2130
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Oferta :"
         Height          =   195
         Left            =   -74910
         TabIndex        =   215
         Top             =   495
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora Final :"
         Height          =   195
         Index           =   24
         Left            =   -74595
         TabIndex        =   214
         Top             =   1725
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora Inicial :"
         Height          =   195
         Index           =   23
         Left            =   -74670
         TabIndex        =   213
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   22
         Left            =   -74340
         TabIndex        =   212
         Top             =   870
         Width           =   540
      End
      Begin VB.Label lblReceta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Height          =   315
         Left            =   -74910
         TabIndex        =   201
         Top             =   2430
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Label lblDirecto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -74910
         TabIndex        =   200
         Top             =   1485
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descargo Receta :"
         Height          =   195
         Index           =   14
         Left            =   -74895
         TabIndex        =   199
         Top             =   2055
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Descargo :"
         Height          =   195
         Index           =   5
         Left            =   -74685
         TabIndex        =   198
         Top             =   555
         Width           =   1140
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descargo Directo :"
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   197
         Top             =   1110
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Area de Impresión :"
         Height          =   195
         Index           =   20
         Left            =   -74640
         TabIndex        =   196
         Top             =   585
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3960
      Left            =   2400
      TabIndex        =   222
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6985
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos Primarios"
      TabPicture(0)   =   "frmProductoDetalle.frx":107C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtTipoProducto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblUsuario"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFecha"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label(32)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label(31)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label(59)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblProductoSunat"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label(63)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblSunatprod"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label(66)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboUnidadNegocio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboGrupo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboArea"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboSubGrupo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdBuscaTipo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCodigo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtResumido"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDetallado"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdRefrescar"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkImprime"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkActivo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdProductoSunat"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txturlimage"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmProductoDetalle.frx":107E4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label(21)"
      Tab(1).Control(1)=   "Label(19)"
      Tab(1).Control(2)=   "Label(61)"
      Tab(1).Control(3)=   "Label(62)"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "imgFoto"
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(7)=   "Label(64)"
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(9)=   "Label(65)"
      Tab(1).Control(10)=   "dlgFoto"
      Tab(1).Control(11)=   "chkMultiArea"
      Tab(1).Control(12)=   "chkDescuento"
      Tab(1).Control(13)=   "chkPropiedad"
      Tab(1).Control(14)=   "chkModificable"
      Tab(1).Control(15)=   "txtBarra"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtAlternativa"
      Tab(1).Control(17)=   "txtTiempo"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkBalanza"
      Tab(1).Control(19)=   "chkLCombinacion"
      Tab(1).Control(20)=   "cmdCombinacion"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtCantidadItems"
      Tab(1).Control(22)=   "cmdAgregarFoto"
      Tab(1).Control(23)=   "cmdImprimeVineta"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chkSolicitaAutorizacionEnvio"
      Tab(1).Control(25)=   "chkDefineArea"
      Tab(1).Control(26)=   "txtPrecioReferencial"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtCantidadItemsMinima"
      Tab(1).Control(28)=   "txtSKU"
      Tab(1).Control(29)=   "chkActivaUsoExterno"
      Tab(1).Control(30)=   "chklComboVisible"
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Leyes aplicables"
      TabPicture(2)   =   "frmProductoDetalle.frx":10800
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "ChtAplicaInafectacion"
      Tab(2).Control(2)=   "ChtAplicaExoneracion"
      Tab(2).Control(3)=   "chkLey1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Integraciones"
      TabPicture(3)   =   "frmProductoDetalle.frx":1081C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chklComboVisible 
         Alignment       =   1  'Right Justify
         Caption         =   "Combo Visible"
         Height          =   255
         Left            =   -71220
         TabIndex        =   293
         Top             =   2190
         Width           =   1335
      End
      Begin VB.TextBox txturlimage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1850
         MaxLength       =   400
         TabIndex        =   261
         Top             =   3360
         Width           =   7440
      End
      Begin VB.CheckBox chkActivaUsoExterno 
         Caption         =   "Activa.(Integracion Uber)"
         Height          =   195
         Left            =   -74775
         TabIndex        =   260
         Top             =   3525
         Width           =   2190
      End
      Begin VB.CheckBox chkLey1 
         Caption         =   "Aplica Ley1"
         Height          =   195
         Left            =   -74760
         TabIndex        =   259
         Top             =   480
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtSKU 
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
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   258
         Top             =   1060
         Width           =   3270
      End
      Begin VB.TextBox txtCantidadItemsMinima 
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
         Left            =   -70080
         MaxLength       =   24
         TabIndex        =   257
         Text            =   " "
         Top             =   1890
         Width           =   480
      End
      Begin VB.TextBox txtPrecioReferencial 
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
         Height          =   315
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   256
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1690
         Width           =   1185
      End
      Begin VB.CommandButton cmdProductoSunat 
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
         Left            =   8520
         Picture         =   "frmProductoDetalle.frx":10838
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   2160
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CheckBox chkDefineArea 
         Caption         =   "Definir Área de Producción al Enviar"
         Height          =   225
         Left            =   -74775
         TabIndex        =   254
         Top             =   3285
         Width           =   2895
      End
      Begin VB.CheckBox chkSolicitaAutorizacionEnvio 
         Caption         =   "Solicita autorización al enviar a producción"
         Height          =   195
         Left            =   -74775
         TabIndex        =   253
         Top             =   3060
         Width           =   4385
      End
      Begin VB.CommandButton cmdImprimeVineta 
         Caption         =   "Viñeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69600
         Picture         =   "frmProductoDetalle.frx":1093A
         Style           =   1  'Graphical
         TabIndex        =   252
         TabStop         =   0   'False
         ToolTipText     =   "Genera Viñeta"
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton cmdAgregarFoto 
         Caption         =   "Editar"
         Height          =   350
         Left            =   -68040
         TabIndex        =   251
         Top             =   2950
         Width           =   1695
      End
      Begin VB.TextBox txtCantidadItems 
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
         Left            =   -70080
         MaxLength       =   24
         TabIndex        =   250
         Text            =   " "
         Top             =   1560
         Width           =   480
      End
      Begin VB.CommandButton cmdCombinacion 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -69600
         Picture         =   "frmProductoDetalle.frx":10D7C
         Style           =   1  'Graphical
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   1560
         Width           =   945
      End
      Begin VB.CheckBox chkLCombinacion 
         Caption         =   "Combinación :"
         Height          =   225
         Left            =   -72930
         TabIndex        =   248
         Top             =   2295
         Width           =   1365
      End
      Begin VB.CheckBox chkBalanza 
         Caption         =   "Ingreso de peso con Balanza"
         Height          =   195
         Left            =   -72930
         TabIndex        =   247
         Top             =   2775
         Width           =   2445
      End
      Begin VB.TextBox txtTiempo 
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
         Height          =   315
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   246
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1380
         Width           =   1185
      End
      Begin VB.TextBox txtAlternativa 
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
         Left            =   -72960
         MaxLength       =   24
         TabIndex        =   245
         Top             =   720
         Width           =   3270
      End
      Begin VB.TextBox txtBarra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   360
         Width           =   3270
      End
      Begin VB.CheckBox chkModificable 
         Caption         =   "Precio Modificable"
         Height          =   225
         Left            =   -74775
         TabIndex        =   243
         Top             =   2790
         Width           =   1695
      End
      Begin VB.CheckBox chkPropiedad 
         Caption         =   "Propiedad Visible"
         Height          =   225
         Left            =   -74775
         TabIndex        =   242
         Top             =   2565
         Width           =   1605
      End
      Begin VB.CheckBox chkDescuento 
         Caption         =   "Permite Descuentos"
         Height          =   195
         Left            =   -74775
         TabIndex        =   241
         Top             =   2325
         Width           =   1845
      End
      Begin VB.CheckBox chkMultiArea 
         Caption         =   "Multi Area de Producción"
         Height          =   225
         Left            =   -72930
         TabIndex        =   240
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   3840
         TabIndex        =   239
         Top             =   2680
         Width           =   855
      End
      Begin VB.CheckBox chkImprime 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime en Area :"
         Height          =   225
         Left            =   525
         TabIndex        =   238
         Top             =   2664
         Width           =   1530
      End
      Begin VB.CommandButton cmdRefrescar 
         Height          =   285
         Left            =   4755
         Picture         =   "frmProductoDetalle.frx":11086
         Style           =   1  'Graphical
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   1178
         Width           =   465
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
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   236
         Top             =   829
         Width           =   7320
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   1860
         MaxLength       =   24
         TabIndex        =   235
         Top             =   1178
         Width           =   2835
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   480
         Width           =   1170
      End
      Begin VB.CommandButton cmdBuscaTipo 
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
         Left            =   8520
         Picture         =   "frmProductoDetalle.frx":111D0
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   1800
         Width           =   675
      End
      Begin VB.Frame Frame6 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   225
         Top             =   360
         Width           =   4335
         Begin VB.CheckBox chkIntegracionWebApi 
            Caption         =   "Integracion Web Api"
            Height          =   315
            Left            =   120
            TabIndex        =   294
            Top             =   2950
            Width           =   2505
         End
         Begin VB.CheckBox chkActivaPatioComida 
            Caption         =   "Activar en Integración"
            Height          =   195
            Left            =   120
            TabIndex        =   229
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtDesPatio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   120
            MaxLength       =   400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   228
            Top             =   1560
            Width           =   4080
         End
         Begin VB.TextBox txtStockPlatoCritico 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   227
            Text            =   "0"
            Top             =   2670
            Width           =   975
         End
         Begin VB.TextBox txtNombrePatio 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   226
            Top             =   840
            Width           =   4095
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Producto:"
            Height          =   195
            Index           =   67
            Left            =   120
            TabIndex        =   232
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Stock plato critico por dìa (0=Ilimitado):"
            Height          =   195
            Index           =   68
            Left            =   120
            TabIndex        =   231
            Top             =   2690
            Width           =   2730
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de producto:"
            Height          =   195
            Index           =   69
            Left            =   120
            TabIndex        =   230
            Top             =   600
            Width           =   1500
         End
      End
      Begin VB.CheckBox ChtAplicaExoneracion 
         Caption         =   "Aplica Exoneración"
         Height          =   255
         Left            =   -74760
         TabIndex        =   224
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox ChtAplicaInafectacion 
         Caption         =   "Aplica Inafectación"
         Height          =   255
         Left            =   -74760
         TabIndex        =   223
         Top             =   1680
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   6255
         TabIndex        =   262
         Top             =   1320
         Width           =   2925
         _ExtentX        =   5159
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1860
         TabIndex        =   263
         Top             =   1906
         Width           =   2835
         _ExtentX        =   5001
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
         Left            =   1860
         TabIndex        =   264
         Top             =   1527
         Width           =   2835
         _ExtentX        =   5001
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
      Begin MSDataListLib.DataCombo cboUnidadNegocio 
         Height          =   315
         Left            =   1860
         TabIndex        =   265
         Top             =   2285
         Width           =   2835
         _ExtentX        =   5001
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
      Begin MSComDlg.CommonDialog dlgFoto 
         Left            =   -67800
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Url Imagen:"
         Height          =   195
         Index           =   66
         Left            =   960
         TabIndex        =   292
         Top             =   3380
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "SKU :"
         Height          =   195
         Index           =   65
         Left            =   -73500
         TabIndex        =   291
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Minima :"
         Height          =   195
         Left            =   -71400
         TabIndex        =   290
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Referencial (Bonificación):"
         Height          =   390
         Index           =   64
         Left            =   -74640
         TabIndex        =   289
         Top             =   1785
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSunatprod 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6240
         TabIndex        =   288
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Producto Sunat:"
         Height          =   435
         Index           =   63
         Left            =   4800
         TabIndex        =   287
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblProductoSunat 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6255
         TabIndex        =   286
         Top             =   2160
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "(Mobile Inforest)"
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
         Height          =   240
         Left            =   -71850
         TabIndex        =   285
         Top             =   3285
         Width           =   1815
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2600
         Left            =   -68640
         Stretch         =   -1  'True
         ToolTipText     =   "Foto"
         Top             =   360
         Width           =   2900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Máxima:"
         Height          =   195
         Left            =   -71400
         TabIndex        =   284
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio :"
         Height          =   195
         Index           =   59
         Left            =   300
         TabIndex        =   283
         Top             =   2345
         Width           =   1470
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "min"
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
         Index           =   62
         Left            =   -71760
         TabIndex        =   282
         Top             =   1380
         Width           =   300
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tiempo Ideal de Preparación :"
         Height          =   390
         Index           =   61
         Left            =   -74760
         TabIndex        =   281
         Top             =   1330
         Width           =   1665
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Alternativa :"
         Height          =   195
         Index           =   19
         Left            =   -74805
         TabIndex        =   280
         Top             =   800
         Width           =   1725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de Barra :"
         Height          =   195
         Index           =   21
         Left            =   -74325
         TabIndex        =   279
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Modificación :"
         Height          =   195
         Index           =   31
         Left            =   300
         TabIndex        =   278
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Modificación :"
         Height          =   195
         Index           =   32
         Left            =   5760
         TabIndex        =   277
         Top             =   2985
         Width           =   1575
      End
      Begin VB.Label lblFecha 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1845
         TabIndex        =   276
         Top             =   3000
         Width           =   2925
      End
      Begin VB.Label lblUsuario 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7440
         TabIndex        =   275
         Top             =   2925
         Width           =   1845
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Area Producción :"
         Height          =   195
         Index           =   7
         Left            =   495
         TabIndex        =   274
         Top             =   1966
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "SubGrupo :"
         Height          =   195
         Index           =   6
         Left            =   5385
         TabIndex        =   273
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   272
         Top             =   874
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   271
         Top             =   1223
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1185
         TabIndex        =   270
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto :"
         Height          =   195
         Index           =   4
         Left            =   4875
         TabIndex        =   269
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Grupo :"
         Height          =   195
         Index           =   3
         Left            =   1275
         TabIndex        =   268
         Top             =   1587
         Width           =   525
      End
      Begin VB.Label txtTipoProducto 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6255
         TabIndex        =   267
         Top             =   1800
         Width           =   2160
      End
      Begin VB.Label Label4 
         Caption         =   "Operación a la que es afecto el producto de venta:"
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
         Left            =   -74760
         TabIndex        =   266
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cargo Automático"
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmProductoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public SAPVALOR As Boolean
Dim rsSAPDatos As Recordset
Dim sCodSapInf As String
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

Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsArea As Recordset
Dim RsPrecio As Recordset
Dim RsDescargo As Recordset
Dim RsBoton As Recordset
Dim RsReceta As Recordset
Dim RsEnlace As Recordset
Dim RsMoneda As Recordset
Dim RsImpresion As Recordset
Dim RsGImpresion As Recordset
Dim rsCombo As Recordset
Dim RsOferta As Recordset
Dim RsFrecuencia As Recordset
Dim RsUnidadNegocio As Recordset

Dim nPrecioVenta As Double
Dim nPrecioDelivery As Double
Dim nPrecioLlevar As Double
Dim nPrecioCanal4 As Double
Dim nPrecioCanal5 As Double
Dim nPrecio As Double
Dim nNetoLlevar As Double
Dim nNetoDelivery As Double
Dim nNetoCanal4 As Double
Dim nNetoCanal5 As Double

Dim nInsumo As Double
Dim nInsumo2 As Double
Dim nInsumo3 As Double
Dim nInsumo4 As Double
Dim nInsumo5 As Double
Dim nGasto As Double
Dim nGasto2 As Double
Dim nGasto3 As Double
Dim nGasto4 As Double
Dim nGasto5 As Double
Dim nMObra As Double
Dim nMObra2 As Double
Dim nMObra3 As Double
Dim nMObra4 As Double
Dim nMObra5 As Double
Dim nCosto As Double
Dim nCosto2 As Double
Dim nCosto3 As Double
Dim nCosto4 As Double
Dim nCosto5 As Double
Dim nUtilidad As Double
Dim nUtilidad2 As Double
Dim nUtilidad3 As Double
Dim nUtilidad4 As Double
Dim nUtilidad5 As Double

Dim sCodInfhotel As String
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nImpuesto4 As Double
Dim nImpuesto5 As Double
Dim nImpuesto6 As Double
Dim nImpuesto7 As Double
Dim nImpuesto8 As Double
Dim nImpuesto9 As Double
Dim nImpuesto10 As Double
Dim nImpuesto11 As Double
Dim nImpuesto12 As Double
Dim nImpuesto13 As Double
Dim nImpuesto14 As Double
Dim nImpuesto15 As Double
Dim i As Integer
Dim sTipoProducto As String

Dim nExiste As Integer

Dim scodigoInsumoCritico As String

Private strFilenameRuta As String
Sub LlenaCombos()
    With cboArea
         Isql = "Select * from vArea order by Descripcion"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboImpresion
         Isql = "Select * from vArea where lActivo= 1 order by Descripcion"
         Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsImpresion
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboGrupo
         Isql = "Select * from vGrupo where lActivo= 1 order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboSubGrupo
         Isql = "Select * from vSubGrupo where lActivo= 1 order by Descripcion"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboDescargo
         Isql = "Select * from vTipoDescargo where lActivo= 1 order by Descripcion"
         Set RsDescargo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsDescargo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboMoneda
         Isql = "Select * from vMoneda  order by Descripcion"
         Set RsMoneda = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMoneda
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboMonedaDelivery
         Set .RowSource = RsMoneda
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    
    With cboMonedaLlevar
         Set .RowSource = RsMoneda
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboMonedaCanal4
         Set .RowSource = RsMoneda
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboMonedaCanal5
         Set .RowSource = RsMoneda
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
           
    With cboUnidadNegocio
         Isql = "Select * from vUnidadNegocio where lActivo= 1"
         Set RsUnidadNegocio = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUnidadNegocio
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
        
End Sub

Sub Asignar()
    With frmProducto.RsCabecera
        
        ' verifica que tenga el check de los precios de venta seran los de almacen
        If Calcular("select lTransVen as codigo from vsubgrupo where codigo =(select tSubGrupo from TPRODUCTO where tDescargo='D' and  tCodigoProducto='" & !codigo & "')", Cn) = True Then
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
        Else
            Frame1.Enabled = True
            Frame2.Enabled = True
            Frame3.Enabled = True
            Frame4.Enabled = True
            Frame5.Enabled = True
        End If
        
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!descripcion), "", !descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtAlternativa = IIf(IsNull(!talternativa), "", !talternativa)
        txtBarra = IIf(IsNull(!tBarra), "", !tBarra)
        lblUsuario.Caption = IIf(IsNull(!tUsuario), "", !tUsuario)
        lblFecha.Caption = IIf(IsNull(!fRegistro), "", !fRegistro)
        chkImpuesto1.Value = IIf(IsNull(!lImpuesto1), 0, IIf(!lImpuesto1, 1, 0))
        chkImpuesto2.Value = IIf(IsNull(!lImpuesto2), 0, IIf(!lImpuesto2, 1, 0))
        chkImpuesto3.Value = IIf(IsNull(!lImpuesto3), 0, IIf(!lImpuesto3, 1, 0))
        chkBalanza.Value = IIf(IsNull(!lBalanza), 0, IIf(!lBalanza, 1, 0))
          If lInfhotel Then
            cmdBuscaTipo.Enabled = False
        End If
        nPrecioVenta = IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta)
        nPrecioDelivery = IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery)
        nPrecioLlevar = IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar)
        nPrecioCanal4 = IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4)
        nPrecioCanal5 = IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5)
        cboMoneda.BoundText = IIf(IsNull(!tMoneda), "", Trim(!tMoneda))

        Select Case pais 'ok
            Case "001" 'Bolivia
                    nPrecio = IIf(nPrecioVenta = 0, 0, nPrecioVenta - (nPrecioVenta * (IIf(chkImpuesto1.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3.Value, nPorcentaje3 / 100, 0))))
                    nNetoDelivery = IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery - (nPrecioDelivery * (IIf(chkImpuesto1Delivery.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Delivery.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Delivery.Value, nPorcentaje3 / 100, 0))))
                    nNetoLlevar = IIf(nPrecioLlevar = 0, nPrecio, nPrecioLlevar - (nPrecioLlevar * (IIf(chkImpuesto1Llevar.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Llevar.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Llevar.Value, nPorcentaje3 / 100, 0))))
                    nNetoCanal4 = IIf(nPrecioCanal4 = 0, nPrecio, nPrecioCanal4 - (nPrecioCanal4 * (IIf(chkImpuesto1Canal4.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal4.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal4.Value, nPorcentaje3 / 100, 0))))
                    nNetoCanal5 = IIf(nPrecioCanal5 = 0, nPrecio, nPrecioCanal5 - (nPrecioCanal5 * (IIf(chkImpuesto1Canal5.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal5.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal5.Value, nPorcentaje3 / 100, 0))))
            Case "003" ' Argentina
                    cboMoneda.BoundText = "01"
                    cboMoneda.Enabled = False
                    nPrecio = IIf(nPrecioVenta = 0, 0, nPrecioVenta / (1 + IIf(chkImpuesto1.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3.Value, nPorcentaje3 / 100, 0)))
                    nNetoDelivery = IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery / (1 + IIf(chkImpuesto1Delivery.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Delivery.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Delivery.Value, nPorcentaje3 / 100, 0)))
                    nNetoLlevar = IIf(nPrecioLlevar = 0, nPrecio, nPrecioLlevar / (1 + IIf(chkImpuesto1Llevar.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Llevar.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Llevar.Value, nPorcentaje3 / 100, 0)))
                    nNetoCanal4 = IIf(nPrecioCanal4 = 0, nPrecio, nPrecioCanal4 / (1 + IIf(chkImpuesto1Canal4.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal4.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal4.Value, nPorcentaje3 / 100, 0)))
                    nNetoCanal5 = IIf(nPrecioCanal5 = 0, nPrecio, nPrecioCanal5 / (1 + IIf(chkImpuesto1Canal5.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal5.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal5.Value, nPorcentaje3 / 100, 0)))
            
            Case Else 'Peru, Ecuador
                    nPrecio = IIf(nPrecioVenta = 0, 0, nPrecioVenta / (1 + IIf(chkImpuesto1.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3.Value, nPorcentaje3 / 100, 0)))
                    nNetoDelivery = IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery / (1 + IIf(chkImpuesto1Delivery.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Delivery.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Delivery.Value, nPorcentaje3 / 100, 0)))
                    nNetoLlevar = IIf(nPrecioLlevar = 0, nPrecio, nPrecioLlevar / (1 + IIf(chkImpuesto1Llevar.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Llevar.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Llevar.Value, nPorcentaje3 / 100, 0)))
                    nNetoCanal4 = IIf(nPrecioCanal4 = 0, nPrecio, nPrecioCanal4 / (1 + IIf(chkImpuesto1Canal4.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal4.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal4.Value, nPorcentaje3 / 100, 0)))
                    nNetoCanal5 = IIf(nPrecioCanal5 = 0, nPrecio, nPrecioCanal5 / (1 + IIf(chkImpuesto1Canal5.Value, nPorcentaje1 / 100, 0) + IIf(chkImpuesto2Canal5.Value, nPorcentaje2 / 100, 0) + IIf(chkImpuesto3Canal5.Value, nPorcentaje3 / 100, 0)))
        End Select
        
        txtPrecioBase.Caption = Format(Round(nPrecio, 2), "###,##0.00")
        txtPrecioVenta.Text = Format(nPrecioVenta, "###,###,##0." & Mid("0000000000", 1, nDecimal))
        sCodInfhotel = IIf(IsNull(!tInfhotel), "", !tInfhotel)
        sTipoProducto = IIf(IsNull(!tTipoProducto), "", Trim(!tTipoProducto))
        txtTipoProducto.Caption = Calcular("select Descripcion as Codigo from vTipoProducto where Codigo='" & sTipoProducto & "'", Cn)
        
        Me.lblProductoSunat.Caption = IIf(IsNull(!tCodigoProductoSunat), "", Trim(!tCodigoProductoSunat))
        Me.txtPrecioReferencial.Text = IIf(IsNull(!nPrecioVentaRefer), "", Trim(!nPrecioVentaRefer))
        
        Me.lblSunatprod.Caption = Calcular("select isnull(Descripcion,'') as codigo from tproductosunat where tcodigoproducto='" & Me.lblProductoSunat.Caption & "'", Cn)
        If IsNull(!lcontrolinsumocritico) = False Then
            If !lcontrolinsumocritico = True Then
                 chkInsumoCritico.Value = 1
            Else
                 chkInsumoCritico.Value = 0
            End If
        Else
            chkInsumoCritico.Value = 0
        End If
        
         If chkInsumoCritico.Value = 1 Then
                    fraInsumo.Enabled = True
                    scodigoInsumoCritico = IIf(IsNull(!tcodigoinsumo), "", Trim(!tcodigoinsumo))
                    lblInsumo.Caption = Calcular("select descripcion as codigo from tinsumo where tcodigo='" & scodigoInsumoCritico & "'", Cn)
        Else
                    fraInsumo.Enabled = False
                    scodigoInsumoCritico = ""
                    lblInsumo.Caption = ""
        End If
    
    
        'Data Combo
        cboArea.BoundText = IIf(IsNull(!tArea), "", Trim(!tArea))
        cboGrupo.BoundText = IIf(IsNull(!tgrupo), "", Trim(!tgrupo))
        RsSubGrupo.Filter = "tGrupo = '" & cboGrupo.BoundText & "'"
        cboSubGrupo.BoundText = IIf(IsNull(!tSubGrupo), "", Trim(!tSubGrupo))
        cboDescargo.BoundText = IIf(IsNull(!tDescargo), "", Trim(!tDescargo))
        cboUnidadNegocio.BoundText = IIf(IsNull(!tUnidadNegocio), "", Trim(!tUnidadNegocio))
        
        'Area de Impresion
        RsGImpresion.Filter = "tCodigoProducto='" & txtCodigo.Text & "'"
        If Not RsGImpresion.EOF Then
           RsGImpresion.MoveFirst
           cboImpresion.BoundText = RsGImpresion!tArea
        Else
           cboImpresion.BoundText = ""
        End If
        txtTiempo.Text = IIf(IsNull(!nTiempo), 0, !nTiempo)
        
        'Combos
        rsCombo.Filter = "tCombo='" & txtCodigo.Text & "'"
        
        ' nuevo
        chkLCombinacion = IIf(!lCombinacion = True, 1, 0)
        If chkLCombinacion.Value = 1 Then
            cmdCombinacion.Enabled = True
        Else
            cmdCombinacion.Enabled = False
        End If
        
        txtCantidadItems.Text = Format(IIf(IsNull(!nCombinacion), 0, !nCombinacion), "##0")
        txtCantidadItemsMinima.Text = Format(IIf(IsNull(!nCombinacionminima), 0, !nCombinacionminima), "##0")
        'Ofertas
        txtNombre.Locked = True
        RsOferta.Filter = "tCodigoProducto='" & txtCodigo.Text & "'"
        If Not RsOferta.EOF Then
           RsOferta.MoveFirst
           AsignaDatos
        Else
           txtNombre.Text = ""
           txtRatio.Text = "0.00"
           ChkActv = False
           dtpFecha.Value = Date
           dtpHoraIni.Value = 0
           dtpHoraFin.Value = 0
           chkAcumulable = False
        End If
                        
        'Costos
        nInsumo = IIf(IsNull(frmProducto.RsCabecera!nInsumo), 0, frmProducto.RsCabecera!nInsumo)
        nInsumo2 = IIf(IsNull(frmProducto.RsCabecera!nInsumo2), 0, frmProducto.RsCabecera!nInsumo2)
        nInsumo3 = IIf(IsNull(frmProducto.RsCabecera!nInsumo3), 0, frmProducto.RsCabecera!nInsumo3)
        nInsumo4 = IIf(IsNull(frmProducto.RsCabecera!nInsumo4), 0, frmProducto.RsCabecera!nInsumo4)
        nInsumo5 = IIf(IsNull(frmProducto.RsCabecera!nInsumo5), 0, frmProducto.RsCabecera!nInsumo5)
        nGasto = IIf(IsNull(frmProducto.RsCabecera!nGasto), 0, frmProducto.RsCabecera!nGasto)
        nGasto2 = IIf(IsNull(frmProducto.RsCabecera!nGasto2), 0, frmProducto.RsCabecera!nGasto2)
        nGasto3 = IIf(IsNull(frmProducto.RsCabecera!nGasto3), 0, frmProducto.RsCabecera!nGasto3)
        nGasto4 = IIf(IsNull(frmProducto.RsCabecera!nGasto4), 0, frmProducto.RsCabecera!nGasto4)
        nGasto5 = IIf(IsNull(frmProducto.RsCabecera!nGasto5), 0, frmProducto.RsCabecera!nGasto5)
        nMObra = IIf(IsNull(frmProducto.RsCabecera!nManoObra), 0, frmProducto.RsCabecera!nManoObra)
        nMObra2 = IIf(IsNull(frmProducto.RsCabecera!nManoObra2), 0, frmProducto.RsCabecera!nManoObra2)
        nMObra3 = IIf(IsNull(frmProducto.RsCabecera!nManoObra3), 0, frmProducto.RsCabecera!nManoObra3)
        nMObra4 = IIf(IsNull(frmProducto.RsCabecera!nManoObra4), 0, frmProducto.RsCabecera!nManoObra4)
        nMObra5 = IIf(IsNull(frmProducto.RsCabecera!nManoObra5), 0, frmProducto.RsCabecera!nManoObra5)
        nCosto = nInsumo + nGasto + nMObra
        nCosto2 = nInsumo2 + nGasto2 + nMObra2
        nCosto3 = nInsumo3 + nGasto3 + nMObra3
        nCosto4 = nInsumo4 + nGasto4 + nMObra4
        nCosto5 = nInsumo5 + nGasto5 + nMObra5
        nUtilidad = IIf(cboMoneda.BoundText = "01", nPrecio, nPrecio * nTC) - nCosto
        nUtilidad2 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery), IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery) * nTC) - nCosto2
        nUtilidad3 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioLlevar = 0, nPrecio, nPrecioLlevar), IIf(nPrecioLlevar = 0, nPrecio, nPrecioLlevar) * nTC) - nCosto3
        nUtilidad4 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioCanal4 = 0, nPrecio, nPrecioCanal4), IIf(nPrecioCanal4 = 0, nPrecio, nPrecioCanal4) * nTC) - nCosto4
        nUtilidad5 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioCanal5 = 0, nPrecio, nPrecioCanal5), IIf(nPrecioCanal5 = 0, nPrecio, nPrecioCanal5) * nTC) - nCosto5
        
        'Foto
        'On Error Resume Next
        'imgFoto.Picture = LoadPicture("bmps\" & txtCodigo & ".BMP")
        'If err.Number = 53 Then
        '   imgFoto.Picture = LoadPicture("bmps\SinFoto.BMP")
        'End If
        
        Grafico
        cmdOpcion(5).Enabled = False
        
        'Descargos
        If lAlmacen Then
           Select Case cboDescargo.BoundText
                  Case Is = "D"
                       cmdBusca(0).Enabled = True
                       txtDirecto.backColor = &H80000005
                       txtDirecto.Enabled = True
                       
                       cmdBusca(1).Enabled = False
                       txtReceta.backColor = &H80000016
                       txtReceta.Text = ""
                       txtReceta.Enabled = False
                       lblReceta.Caption = ""
                       txtDirecto.Text = IIf(IsNull(!tEnlace), "", Trim(!tEnlace))
                       lblDirecto.Caption = Calcular("select tDetallado as Codigo from TPRODUCTO where tCodigoProducto='" & txtDirecto.Text & "'", CnAlmacen)
                       lblDirecto.Caption = IIf(lblDirecto.Caption = "0", "", lblDirecto.Caption)
                       
                  Case Is = "R"
                       cmdOpcion(5).Enabled = True
                       cmdBusca(1).Enabled = True
                       txtReceta.backColor = &H80000005
                       txtReceta.Enabled = True
                       
                       cmdBusca(0).Enabled = False
                       txtDirecto.backColor = &H80000016
                       txtDirecto.Enabled = False
                       txtDirecto.Text = ""
                       lblDirecto.Caption = ""
                       txtReceta.Text = IIf(IsNull(!tEnlace), "", Trim(!tEnlace))
                       lblReceta.Caption = Calcular("select tDescripcion as Codigo from MRECETAVENTA where tRecetaVenta='" & txtReceta.Text & "' and tLocal='" & sLocal & "'", CnAlmacen)
                       lblReceta.Caption = IIf(lblReceta.Caption = "0", "", lblReceta.Caption)

                  Case Else
                       cmdBusca(1).Enabled = False
                       txtReceta.backColor = &H80000016
                       txtReceta.Text = ""
                       lblReceta.Caption = ""
                       txtReceta.Enabled = False
                       
                       cmdBusca(0).Enabled = False
                       txtDirecto.backColor = &H80000016
                       txtDirecto.Text = ""
                       txtDirecto.Enabled = False
                       lblDirecto.Caption = ""
           End Select
        End If
                                
        'Check Box
        chkImprime = IIf(!lImprimearea = True, 1, 0)
        chkModificable = IIf(!lModificable = True, 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkDescuento = IIf(!lDescuento = True, 1, 0)
        chkPropiedad = IIf(!lPropiedad = True, 1, 0)
        
        chkSolicitaAutorizacionEnvio = IIf(!lSolicitaAutorizacionEnvio = True, 1, 0)
        
        'multiarea
        If IsNull(!lmultiarea) = True Then
            chkMultiArea.Value = 0
        ElseIf (!lmultiarea) = True Then
                chkMultiArea.Value = 1
        ElseIf (!lmultiarea) = False Then
                chkMultiArea.Value = 0
        End If
                       
        chkVisualizarLocal.Value = IIf(IsNull(!lLocal), 0, IIf(!lLocal, 1, 0))
        chkVisualizarDelivery.Value = IIf(IsNull(!lDelivery), 0, IIf(!lDelivery, 1, 0))
        chkVisualizarLlevar.Value = IIf(IsNull(!lLlevar), 0, IIf(!lLlevar, 1, 0))
        chkvisualizarCanal4.Value = IIf(IsNull(!lCanal4), 0, IIf(!lCanal4, 1, 0))
        chkVisualizarCanal5.Value = IIf(IsNull(!lCanal5), 0, IIf(!lCanal5, 1, 0))
                
        Me.chkDefineArea.Value = IIf(IsNull(!lDefineAreaProduccion), 0, IIf(!lDefineAreaProduccion, 1, 0))
        Me.txtSKU.Text = IIf(IsNull(!tCodigoSKU), "", !tCodigoSKU)
                
        'Precios por tipo de pedido
        cboMonedaLlevar.BoundText = cboMoneda.BoundText
        cboMonedaDelivery.BoundText = cboMoneda.BoundText
        cboMonedaCanal4.BoundText = cboMoneda.BoundText
        cboMonedaCanal5.BoundText = cboMoneda.BoundText
                
        Me.chkLey1.Value = IIf(IsNull(!lley1), 0, IIf(!lley1, 1, 0))
        Me.chkActivaUsoExterno.Value = IIf(IsNull(!lActivahomologacionexterno), 0, IIf(!lActivahomologacionexterno, 1, 0))
        'GCAA 03022020
        Me.txturlimage.Text = IIf(IsNull(!surlimage), "", !surlimage)
        Me.chkActivaPatioComida.Value = IIf(IsNull(!lActivaPatioComida), 0, IIf(!lActivaPatioComida, 1, 0)) '
        Me.txtDesPatio.Text = IIf(IsNull(!tDescripcionPatio), "", !tDescripcionPatio)
        Me.txtNombrePatio.Text = IIf(IsNull(!tNombrePatio), "", !tNombrePatio)
        Me.txtStockPlatoCritico.Text = IIf(IsNull(!tStockPlatoCritico), "", !tStockPlatoCritico)
        ''' ---- gcaa 19102020
        Me.ChtAplicaExoneracion.Value = IIf(IsNull(!taplicaexoneracion), 0, IIf(!taplicaexoneracion, 1, 0)) '
        Me.ChtAplicaInafectacion.Value = IIf(IsNull(!taplicainafectacion), 0, IIf(!taplicainafectacion, 1, 0)) '

      Me.chklComboVisible.Value = IIf(IsNull(!lComboVisible), 0, IIf(!lComboVisible, 1, 0)) '
  Me.chkIntegracionWebApi.Value = IIf(IsNull(!IntegracionWebApi), 0, IIf(!IntegracionWebApi, 1, 0)) '

 
        
        
        If IsNull(!nPrecioDelivery) Or !nPrecioDelivery = 0 Then
           nPrecioDelivery = Round(nPrecioVenta * (1 + (nDELIVERY / 100)), 2)
           txtDelivery.Text = Format(nPrecioDelivery, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkDelivery.Value = 0
           txtDelivery.Enabled = False
           
           chkImpuesto1Delivery.Value = IIf(!lImpuesto1, 1, 0)
           chkImpuesto2Delivery.Value = IIf(!lImpuesto2, 1, 0)
           chkImpuesto3Delivery.Value = IIf(!lImpuesto3, 1, 0)
        Else
           nPrecioDelivery = !nPrecioDelivery
           txtDelivery.Text = Format(nPrecioDelivery, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkDelivery.Value = 1
           txtDelivery.Enabled = True
                      
           chkImpuesto1Delivery.Value = IIf(!lImpuesto4, 1, 0)
           chkImpuesto2Delivery.Value = IIf(!lImpuesto5, 1, 0)
           chkImpuesto3Delivery.Value = IIf(!lImpuesto6, 1, 0)
        End If
        Calcula2
        
        If IsNull(!nPrecioLlevar) Or !nPrecioLlevar = 0 Then
           nPrecioLlevar = Round(nPrecioVenta * (1 + (nLlevar / 100)), 2)
           txtLlevar.Text = Format(nPrecioLlevar, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkLlevar.Value = 0
           txtLlevar.Enabled = False
           chkImpuesto1Llevar.Value = IIf(!lImpuesto1, 1, 0)
           chkImpuesto2Llevar.Value = IIf(!lImpuesto2, 1, 0)
           chkImpuesto3Llevar.Value = IIf(!lImpuesto3, 1, 0)
        Else
           nPrecioLlevar = !nPrecioLlevar
           txtLlevar.Text = Format(nPrecioLlevar, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkLlevar.Value = 1
           txtLlevar.Enabled = True
           
           chkImpuesto1Llevar.Value = IIf(!lImpuesto7, 1, 0)
           chkImpuesto2Llevar.Value = IIf(!lImpuesto8, 1, 0)
           chkImpuesto3Llevar.Value = IIf(!lImpuesto9, 1, 0)
        End If
        Calcula3
               
        If IsNull(!nPrecioCanal4) Or !nPrecioCanal4 = 0 Then
           nPrecioCanal4 = Round(nPrecioVenta * (1 + (nCanal4 / 100)), 2)
           txtCanal4.Text = Format(nPrecioCanal4, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkCanal4.Value = 0
           txtCanal4.Enabled = False
           chkImpuesto1Canal4.Value = IIf(!lImpuesto10, 1, 0)
           chkImpuesto2Canal4.Value = IIf(!lImpuesto11, 1, 0)
           chkImpuesto3Canal4.Value = IIf(!lImpuesto12, 1, 0)
        Else
           nPrecioCanal4 = !nPrecioCanal4
           txtCanal4.Text = Format(nPrecioCanal4, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkCanal4.Value = 1
           txtCanal4.Enabled = True
           
           chkImpuesto1Canal4.Value = IIf(!lImpuesto10, 1, 0)
           chkImpuesto2Canal4.Value = IIf(!lImpuesto11, 1, 0)
           chkImpuesto3Canal4.Value = IIf(!lImpuesto12, 1, 0)
        End If
        Calcula4
        
        If IsNull(!nPrecioCanal5) Or !nPrecioCanal5 = 0 Then
           nPrecioCanal5 = Round(nPrecioVenta * (1 + (nCanal5 / 100)), 2)
           txtCanal5.Text = Format(nPrecioCanal5, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkCanal5.Value = 0
           txtCanal5.Enabled = False
           chkImpuesto1Canal5.Value = IIf(!lImpuesto13, 1, 0)
           chkImpuesto2Canal5.Value = IIf(!lImpuesto14, 1, 0)
           chkImpuesto3Canal5.Value = IIf(!lImpuesto15, 1, 0)
        Else
           nPrecioCanal5 = !nPrecioCanal5
           txtCanal5.Text = Format(nPrecioCanal5, "###,##0." & Mid("0000000000", 1, nDecimal))
           chkCanal5.Value = 1
           txtCanal5.Enabled = True
           
           chkImpuesto1Canal5.Value = IIf(!lImpuesto13, 1, 0)
           chkImpuesto2Canal5.Value = IIf(!lImpuesto14, 1, 0)
           chkImpuesto3Canal5.Value = IIf(!lImpuesto15, 1, 0)
        End If
        Calcula5
                
        Botonera
    End With
    Calcula
    RsSubGrupo.Filter = "tGrupo = '" & cboGrupo.BoundText & "'"
    
      
       'fotoooooooooo
       On Error GoTo Err
       Dim rst1 As New ADODB.Recordset
       imgFoto.DataField = "foto"
       cmdAgregarFoto.Caption = "Editar"
       Set rst1 = Lib.OpenRecordset("select ofoto as foto from tproducto where tcodigoproducto='" & txtCodigo.Text & "'", Cn)
       Set imgFoto.DataSource = rst1
       Exit Sub
       
Err:
    imgFoto.Picture = Nothing
    
End Sub

Private Sub cboDescargo_Change()

   If lAlmacen Then
      Select Case cboDescargo.BoundText
             Case Is = "D"
                  cmdBusca(0).Enabled = True
                  txtDirecto.backColor = &H80000005
                  txtDirecto.Enabled = True
                  
                  cmdBusca(1).Enabled = False
                  txtReceta.backColor = &H80000016
                  txtReceta.Text = ""
                  txtReceta.Enabled = False
                  lblReceta.Caption = ""
                  
             Case Is = "R"
                  cmdBusca(1).Enabled = True
                  txtReceta.backColor = &H80000005
                  txtReceta.Enabled = True
                  
                  cmdBusca(0).Enabled = False
                  txtDirecto.backColor = &H80000016
                  txtDirecto.Enabled = False
                  txtDirecto.Text = ""
                  lblDirecto.Caption = ""
                               
             Case Else
                  cmdBusca(1).Enabled = False
                  txtReceta.backColor = &H80000016
                  txtReceta.Text = ""
                  lblReceta.Caption = ""
                  txtReceta.Enabled = False
                  
                  cmdBusca(0).Enabled = False
                  txtDirecto.backColor = &H80000016
                  txtDirecto.Text = ""
                  txtDirecto.Enabled = False
                  lblDirecto.Caption = ""
      End Select
   End If
   Grafico
   
End Sub

Private Sub cboGrupo_Change()
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo where lActivo= 1 and tGrupo = '" & cboGrupo.BoundText & "' order by Descripcion "
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
    End With
End Sub

Private Sub cboMoneda_Click(Area As Integer)
   cboMonedaLlevar.BoundText = cboMoneda.BoundText
   cboMonedaDelivery.BoundText = cboMoneda.BoundText
   cboMonedaCanal4.BoundText = cboMoneda.BoundText
   cboMonedaCanal5.BoundText = cboMoneda.BoundText
End Sub

Private Sub cboSubGrupo_Click(Area As Integer)
    Botonera
    If Sw Then
       If RsSubGrupo.RecordCount > 0 Then
          RsSubGrupo.MoveFirst
          RsSubGrupo.Find ("Codigo ='" & cboSubGrupo.BoundText & "'")
          If Not RsSubGrupo.EOF Then
             chkImpuesto1.Value = IIf(RsSubGrupo!lImpuesto1 And sImpuesto1 <> "", 1, 0)
             chkImpuesto2.Value = IIf(RsSubGrupo!lImpuesto2 And sImpuesto2 <> "", 1, 0)
             chkImpuesto3.Value = IIf(RsSubGrupo!lImpuesto3 And sImpuesto3 <> "", 1, 0)
             chkImprime.Value = IIf(RsSubGrupo!lImprimearea, 1, 0)
             cboArea.BoundText = IIf(IsNull(RsSubGrupo!tArea), "", RsSubGrupo!tArea)
          End If
       End If
    End If
End Sub

Private Sub Check1_Click()

End Sub
 
Private Sub chkImpuesto1Delivery_Click()
   If chkImpuesto1Delivery.Value = 0 Then
        chkImpuesto1.Value = 0
        chkImpuesto1Delivery.Value = 0
        chkImpuesto1Llevar.Value = 0
        chkImpuesto1Canal4.Value = 0
        chkImpuesto1Canal5.Value = 0
    Else
        chkImpuesto1.Value = 1
        chkImpuesto1Delivery.Value = 1
        chkImpuesto1Llevar.Value = 1
        chkImpuesto1Canal4.Value = 1
        chkImpuesto1Canal5.Value = 1
   End If
      
   Calcula2
End Sub

Private Sub chkImpuesto1Llevar_Click()
  If chkImpuesto1Llevar.Value = 0 Then
        chkImpuesto1.Value = 0
        chkImpuesto1Delivery.Value = 0
        chkImpuesto1Llevar.Value = 0
        chkImpuesto1Canal4.Value = 0
        chkImpuesto1Canal5.Value = 0
    Else
        chkImpuesto1.Value = 1
        chkImpuesto1Delivery.Value = 1
        chkImpuesto1Llevar.Value = 1
        chkImpuesto1Canal4.Value = 1
        chkImpuesto1Canal5.Value = 1
   End If
   Calcula3
End Sub

Private Sub chkImpuesto2Delivery_Click()
   Calcula2
End Sub

Private Sub chkImpuesto2Llevar_Click()
   Calcula3
End Sub

Private Sub chkImpuesto3Delivery_Click()
   Calcula2
End Sub

Private Sub chkImpuesto3Llevar_Click()
   Calcula3
End Sub

Private Sub chkImpuesto1Canal4_Click()
  If chkImpuesto1Canal4.Value = 0 Then
        chkImpuesto1.Value = 0
        chkImpuesto1Delivery.Value = 0
        chkImpuesto1Llevar.Value = 0
        chkImpuesto1Canal4.Value = 0
        chkImpuesto1Canal5.Value = 0
    Else
        chkImpuesto1.Value = 1
        chkImpuesto1Delivery.Value = 1
        chkImpuesto1Llevar.Value = 1
        chkImpuesto1Canal4.Value = 1
        chkImpuesto1Canal5.Value = 1
   End If
   Calcula4
End Sub

Private Sub chkImpuesto2Canal4_Click()
   Calcula4
End Sub

Private Sub chkImpuesto3Canal4_Click()
   Calcula4
End Sub

Private Sub chkImpuesto1Canal5_Click()
If chkImpuesto1Canal5.Value = 0 Then
        chkImpuesto1.Value = 0
        chkImpuesto1Delivery.Value = 0
        chkImpuesto1Llevar.Value = 0
        chkImpuesto1Canal4.Value = 0
        chkImpuesto1Canal5.Value = 0
    Else
        chkImpuesto1.Value = 1
        chkImpuesto1Delivery.Value = 1
        chkImpuesto1Llevar.Value = 1
        chkImpuesto1Canal4.Value = 1
        chkImpuesto1Canal5.Value = 1
   End If
   Calcula5
End Sub

Private Sub chkImpuesto2Canal5_Click()
   Calcula5
End Sub

Private Sub chkImpuesto3Canal5_Click()
   Calcula5
End Sub

Private Sub chkLCombinacion_Click()
    If chkLCombinacion.Value Then
        tabProducto.TabEnabled(3) = False
        cmdCombinacion.Enabled = True

        rsCombo.Requery
    Else
         tabProducto.TabEnabled(3) = True
         cmdCombinacion.Enabled = False
         Cn.Execute "delete from TCOMBO where tCombo='" & txtCodigo.Text & "'"
    End If
End Sub

Private Sub ChtAplicaExoneracion_Click()

If ChtAplicaExoneracion.Value = 1 Then
    ChtAplicaInafectacion.Value = 0
    chkImpuesto1.Value = False
    chkImpuesto1Delivery.Value = False
    chkImpuesto1Llevar.Value = False
    chkImpuesto1Canal4.Value = False
    chkImpuesto1Canal5.Value = False
    
    chkImpuesto1.Enabled = False
    chkImpuesto1Delivery.Enabled = False
    chkImpuesto1Llevar.Enabled = False
    chkImpuesto1Canal4.Enabled = False
    chkImpuesto1Canal5.Enabled = False
Else
    chkImpuesto1.Value = 1
    chkImpuesto1Delivery.Value = 1
    chkImpuesto1Llevar.Value = 1
    chkImpuesto1Canal4.Value = 1
    chkImpuesto1Canal5.Value = 1
    
    chkImpuesto1.Enabled = True
    chkImpuesto1Delivery.Enabled = True
    chkImpuesto1Llevar.Enabled = True
    chkImpuesto1Canal4.Enabled = True
    chkImpuesto1Canal5.Enabled = True
End If

End Sub

Private Sub ChtAplicaInafectacion_Click()
If ChtAplicaInafectacion.Value = 1 Then
    ChtAplicaExoneracion.Value = 0
    
    chkImpuesto1.Value = False
    chkImpuesto1Delivery.Value = False
    chkImpuesto1Llevar.Value = False
    chkImpuesto1Canal4.Value = False
    chkImpuesto1Canal5.Value = False
    
    
    chkImpuesto1.Enabled = False
    chkImpuesto1Delivery.Enabled = False
    chkImpuesto1Llevar.Enabled = False
    chkImpuesto1Canal4.Enabled = False
    chkImpuesto1Canal5.Enabled = False
Else
    chkImpuesto1.Value = 1
    chkImpuesto1Delivery.Value = 1
    chkImpuesto1Llevar.Value = 1
    chkImpuesto1Canal4.Value = 1
    chkImpuesto1Canal5.Value = 1

    chkImpuesto1.Enabled = True
    chkImpuesto1Delivery.Enabled = True
    chkImpuesto1Llevar.Enabled = True
    chkImpuesto1Canal4.Enabled = True
    chkImpuesto1Canal5.Enabled = True

End If
End Sub


Private Sub cmdAgregarFoto_Click()


'on error GoTo ErrHandler
If txtCodigo.Text <> "" Then
    dlgFoto.CancelError = False
    With cmdAgregarFoto
        If .Caption = "Editar" Then
            dlgFoto.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFoto.FileName = ""
            dlgFoto.ShowOpen
            imgFoto.Visible = True
            If dlgFoto.FileName <> "" Then
                .Caption = "Guardar"
                strFilenameRuta = dlgFoto.FileName
                imgFoto.Picture = LoadPicture(strFilenameRuta)
            End If
        Else
            
            Call GuardarFoto
            .Caption = "Editar"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para el Producto"
    Exit Sub
End If
errHandler:
strFilenameRuta = ""
    Exit Sub
End Sub

Public Sub GuardarFoto()
        
        Dim imgTeacher()      As Byte
        Dim varPhoto          As Variant
        Dim numfile           As Long
        If (strFilenameRuta <> "") Then
            varPhoto = FileLen(strFilenameRuta)
            ReDim bufimages(varPhoto - 1) As Byte
            numfile = FreeFile
            Open strFilenameRuta For Binary As #numfile
            Get #numfile, , bufimages
            Close #numfile
             imgTeacher = bufimages
        End If
        If (strFilenameRuta = "") Then
            imgTeacher = LoadResData(101, "CUSTOM")
            varPhoto = UBound(imgTeacher)
        End If
        Dim lnfoto As Variant
        lnfoto = varPhoto
        Dim Cmd As New ADODB.Command
        Dim prm As New ADODB.Parameter
        With Cmd
                .ActiveConnection = Cn
                .CommandText = "sp_UpdFotoProducto"
                .CommandType = adCmdStoredProc
        End With
        Set prm = Cmd.CreateParameter("@tCodigo", adChar, adParamInput, 10, txtCodigo.Text)
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


 

Private Sub cmdArea_Click(Index As Integer)
   Select Case Index
          Case Is = 1 ' Agregar Impresion
               If cboImpresion.BoundText = "" Then
                  MsgBox "Area de Impresión en blanco", vbCritical, sMensaje
                  Exit Sub
               End If
               
               If Calcular("select tArea as Codigo from TPRODUCTOAREA where tCodigoProducto='" & txtCodigo.Text & "' and tArea = '" & cboImpresion.BoundText & "'", Cn) = "0" Then
                    'KDS
                    If (KDS_ValidarProductoArea(txtCodigo.Text, cboImpresion.BoundText)) Then
                        MsgBox "No se puede elegir dos areas del Tipo KDS", vbCritical, sMensaje
                        Exit Sub
                    End If

                     'impresion en pantablla
'                    If validadProductoenAreaPantalla(txtCodigo.Text, cboImpresion.BoundText) Then
'                        MsgBox "No se puede agregar dos areas del Tipo Impresion en Pantalla", vbCritical, sMensaje
'                        Exit Sub
'                    End If
                    
                  Cn.Execute "insert into TPRODUCTOAREA (tCodigoProducto, tArea,lNuevoModificado) values ('" & txtCodigo.Text & "', '" & cboImpresion.BoundText & "',1)"
                  RsGImpresion.Requery
               Else
                  MsgBox "Area de Impresión ya existe", vbCritical, sMensaje
                  cboImpresion.BoundText = ""
               End If
                        
          Case Is = 2 ' Eliminar Impresion
               Cn.Execute "delete from TPRODUCTOAREA where tCodigoProducto='" & txtCodigo.Text & "' and tArea = '" & grdGrilla.Columns(0).Value & "'"
               RsGImpresion.Requery
               If Not RsGImpresion.EOF Then
                  RsGImpresion.MoveFirst
               End If
   End Select
End Sub

Private Sub cmdBusca_Click(Index As Integer)
   If Index = 0 Then
      Isql = "Select tCodigoProducto as Codigo, tDetallado as Descripcion, UnidadEntrada from vPRODUCTO where lDDirecto=1 and lActivo = 1"
      frmBuscaAlmacen.nPredeterm = 1
      Call ConfGrilla(3, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                             "Producto", 2, "Descripcion", 4100, 0, 0, "", _
                                             "Unidad", 2, "UnidadEntrada", 2000, 0, 0, "")
      frmBuscaAlmacen.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      txtDirecto.Text = sCodigo
      lblDirecto.Caption = sDescrip
      nInsumo = Calcular("select nPrecioPromedio as Codigo from TPRODUCTO where tCodigoProducto='" & sCodigo & "'", CnAlmacen)
      nInsumo2 = nInsumo
      nInsumo3 = nInsumo
      nInsumo4 = nInsumo
      nInsumo5 = nInsumo
      nCosto = nInsumo
      nCosto2 = nInsumo
      nCosto3 = nInsumo
      nCosto4 = nInsumo
      nCosto5 = nInsumo
      nMObra = 0
      nMObra2 = 0
      nMObra3 = 0
      nMObra4 = 0
      nMObra5 = 0
      nGasto = 0
      nGasto2 = 0
      nGasto3 = 0
      nGasto4 = 0
      nGasto5 = 0
      nUtilidad = 0
      nUtilidad2 = 0
      nUtilidad3 = 0
      nUtilidad4 = 0
      nUtilidad5 = 0
   Else
      Isql = "Select tRecetaVenta as Codigo, tDescripcion as Descripcion, Area from vRecetaVenta where lActivo = 1  and tLocal='" & sLocal & "' order by Descripcion"
      frmBuscaAlmacen.nPredeterm = 1
      Call ConfGrilla(3, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4100, 0, 0, "", _
                                                    "Area", 2, "Area", 2000, 0, 0, "")
      frmBuscaAlmacen.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      txtReceta.Text = sCodigo
      lblReceta.Caption = sDescrip
      Dim RsReceta As Recordset
      Set RsReceta = Lib.OpenRecordset("select * from MRECETAVENTA where tRecetaVenta='" & sCodigo & "' and tLocal='" & sLocal & "'", CnAlmacen)
      If Not RsReceta.EOF Then
         nInsumo = IIf(IsNull(RsReceta!nInsumo), 0, RsReceta!nInsumo)
         nInsumo2 = IIf(IsNull(RsReceta!nInsumo2), 0, RsReceta!nInsumo2)
         nInsumo3 = IIf(IsNull(RsReceta!nInsumo3), 0, RsReceta!nInsumo3)
         nInsumo4 = IIf(IsNull(RsReceta!nInsumo4), 0, RsReceta!nInsumo4)
         nInsumo5 = IIf(IsNull(RsReceta!nInsumo5), 0, RsReceta!nInsumo5)
         nGasto = IIf(IsNull(RsReceta!nGastoGeneral), 0, RsReceta!nGastoGeneral)
         nGasto2 = IIf(IsNull(RsReceta!nGastoGeneral2), 0, RsReceta!nGastoGeneral2)
         nGasto3 = IIf(IsNull(RsReceta!nGastoGeneral3), 0, RsReceta!nGastoGeneral3)
         nGasto4 = IIf(IsNull(RsReceta!nGastoGeneral4), 0, RsReceta!nGastoGeneral4)
         nGasto5 = IIf(IsNull(RsReceta!nGastoGeneral5), 0, RsReceta!nGastoGeneral5)
         nMObra = IIf(IsNull(RsReceta!nManoObra), 0, RsReceta!nManoObra)
         nMObra2 = IIf(IsNull(RsReceta!nManoObra2), 0, RsReceta!nManoObra2)
         nMObra3 = IIf(IsNull(RsReceta!nManoObra3), 0, RsReceta!nManoObra3)
         nMObra4 = IIf(IsNull(RsReceta!nManoObra4), 0, RsReceta!nManoObra4)
         nMObra5 = IIf(IsNull(RsReceta!nManoObra5), 0, RsReceta!nManoObra5)
         nCosto = nInsumo + nGasto + nMObra
         nCosto2 = nInsumo2 + nGasto2 + nMObra2
         nCosto3 = nInsumo3 + nGasto3 + nMObra3
         nCosto4 = nInsumo4 + nGasto4 + nMObra4
         nCosto5 = nInsumo5 + nGasto5 + nMObra5
      Else
         nCosto = 0
         nCosto2 = 0
         nCosto3 = 0
         nCosto4 = 0
         nCosto5 = 0
         nInsumo = 0
         nInsumo2 = 0
         nInsumo3 = 0
         nInsumo4 = 0
         nInsumo5 = 0
         nMObra = 0
         nMObra2 = 0
         nMObra3 = 0
         nMObra4 = 0
         nMObra5 = 0
         nGasto = 0
         nGasto2 = 0
         nGasto3 = 0
         nGasto4 = 0
         nGasto5 = 0
         nUtilidad = 0
         nUtilidad2 = 0
         nUtilidad3 = 0
         nUtilidad4 = 0
         nUtilidad5 = 0
      End If
   End If
   Grafico
End Sub

Private Sub cmdBuscaInsumo_Click()
 
    Isql = "select tCodigo as codigo, Descripcion from tInsumo where lActivo=1 order by Descripcion"
    Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                           "Insumo", 2, "Descripcion", 6100, 0, 0, "")
    frmBusca.nPredeterm = 1
    frmBusca.Show vbModal
    If wEnter Then
       scodigoInsumoCritico = sCodigo
       lblInsumo.Caption = sDescrip
    End If
 End Sub

Private Sub cmdBuscaTipo_Click()
    Isql = "select Codigo, Descripcion from vTipoProducto where lActivo=1 order by Descripcion"
    Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                           "Tipo de Producto", 2, "Descripcion", 6100, 0, 0, "")
    frmBusca.nPredeterm = 1
    frmBusca.Show vbModal
    If wEnter Then
       sTipoProducto = sCodigo
       txtTipoProducto.Caption = sDescrip
    End If
End Sub

Private Sub cmdCombinación_Click()
    frmCombosDetalle.Show vbModal
End Sub

Private Sub cmdCombinacion_Click()
    If txtCodigo.Text <> "" Then
        If val(txtCantidadItems.Text) > 0 And val(txtCantidadItemsMinima.Text) > 0 Then
            If val(txtCantidadItemsMinima.Text) > val(txtCantidadItems.Text) Then MsgBox "La cantidad de items combo minima, no puede ser mayor a la cantidad de items combo maxima!!", vbInformation: txtCantidadItemsMinima.Text = txtCantidadItems.Text
            Cn.Execute "update TPRODUCTO set lCombinacion=1, nCombinacion=" & txtCantidadItems & ", ncombinacionminima = " & txtCantidadItemsMinima & " , lreplica=1 where tCodigoProducto='" & txtCodigo.Text & "'"
            frmCombosDetalle.obtieneCodigoProducto txtCodigo.Text
            frmCombosDetalle.lblCantidad.Caption = txtCantidadItems.Text
            frmCombosDetalle.lblCantidadMinima.Caption = txtCantidadItemsMinima.Text
            frmCombosDetalle.Caption = "Combo: " & Trim(txtDetallado.Text)
            frmCombosDetalle.Show vbModal
        Else
            MsgBox "Ingresar una cantidad válida", vbExclamation, sMensaje
            txtCantidadItems.SetFocus
        End If
    Else
         MsgBox "Debe grabar el Producto, para poder agregar los Items", vbExclamation, sMensaje
    End If
End Sub

Private Sub cmdImprimeVineta_Click()
On Error GoTo Err:
        If txtCodigo.Text = "" Then
            Exit Sub
         End If
        If Len(txtBarra.Text) = 0 Then
            Exit Sub
        End If
        If chkActivo.Value = 0 Then
            Exit Sub
        End If
        If Calcular("select count(*) codigo from timpresora where tcaja='" & sCaja & "' and isnull(lvineta,0)=1", Cn) = 0 Then
            MsgBox "No existe Impresora de Viñetas para esta Caja", vbInformation, sMensaje
            Exit Sub
        End If
        
        If FileExists(App.Path & "\GeneracionEtiqueta\Generacionetiqueta.exe") = False Then
            MsgBox "No tiene el Módulo de Impresión de Viñetas " & vbCrLf & "Comuníquese con Infhotel Servicios Informáticos", vbInformation, sMensaje
            Exit Sub
        End If

        'Shell App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe" & " " & "0" + " & txtCodigo.Text & " + " & sCaja, vbNormalFocus"
        'Shell App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe" & " " & "0" + "0000001" + "001", vbNormalFocus
        
        'Shell App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe" & " " & "0" + txtCodigo.Text + sCaja, vbNormalFocus
        ShellExecute Me.hwnd, "open", App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe", "0" + txtCodigo.Text + sCaja, "C:\", SW_SHOWNORMAL
        Sleep 2000
        
        Exit Sub
Err:
        MsgBox "Problema de Impresión " + Err.Description, sMensaje
End Sub

Private Sub cmdOferta_Click()
    Cn.Execute "delete from TOFERTA where tCodigoProducto='" & txtCodigo.Text & "' and toferta = '" & grdOferta.Columns(0).Value & "'"
    
    RsOferta.Requery
    
    If Not RsOferta.EOF Then
       RsOferta.MoveFirst
    End If
    
    MsgBox "Registro Eliminado", vbInformation, "Mensaje del Sistema"
    
End Sub

Private Sub cmdProductoSunat_Click()
    Isql = "select Codigo, Descripcion from vProductoSunat order by Descripcion"
    Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                           "Producto Sunat", 2, "Descripcion", 6100, 0, 0, "")
    frmBusca.nPredeterm = 1
    frmBusca.Show vbModal
    If wEnter Then
       'sTipoProducto = sCodigo
       lblProductoSunat.Caption = sCodigo 'sDescrip
       Me.lblSunatprod.Caption = sDescrip
    End If
End Sub

Private Sub cmdRefrescar_Click()
   If txtBoton.Caption <> "NA" Then
      cmdBoton(val(txtBoton.Caption)).Caption = txtResumido.Text
   End If
End Sub

Private Sub chkInsumoCritico_click()
    If chkInsumoCritico.Value = 0 Then
            fraInsumo.Enabled = False
            scodigoInsumoCritico = ""
            lblInsumo.Caption = ""
    Else
            fraInsumo.Enabled = True
            scodigoInsumoCritico = ""
            lblInsumo.Caption = ""
    End If
End Sub

Private Sub chkDelivery_Click()
   If chkDelivery.Value = 0 Then
      nPrecioDelivery = Round(nPrecioVenta * (1 + (nDELIVERY / 100)), 2)
      txtDelivery.Text = Format(nPrecioDelivery, "###,##0." & Mid("0000000000", 1, nDecimal))
      txtDelivery.Enabled = False
   Else
      txtDelivery.Enabled = True
   End If
End Sub

Private Sub chkImpuesto1_Click()
   If chkImpuesto1.Value = 0 Then
        chkImpuesto1.Value = 0
        chkImpuesto1Delivery.Value = 0
        chkImpuesto1Llevar.Value = 0
        chkImpuesto1Canal4.Value = 0
        chkImpuesto1Canal5.Value = 0
    Else
        chkImpuesto1Delivery.Value = 1
        chkImpuesto1Llevar.Value = 1
        chkImpuesto1Canal4.Value = 1
        chkImpuesto1Canal5.Value = 1
   End If
   
   
   Calcula1
End Sub

Private Sub chkImpuesto2_Click()
   Calcula1
End Sub

Private Sub chkImpuesto3_Click()
   Calcula1
End Sub

Private Sub chkLlevar_Click()
   If chkLlevar.Value = 0 Then
      nPrecioLlevar = Round(nPrecioVenta * (1 + (nLlevar / 100)), 2)
      txtLlevar.Text = Format(nPrecioLlevar, "###,##0." & Mid("0000000000", 1, nDecimal))
      txtLlevar.Enabled = False
   Else
      txtLlevar.Enabled = True
   End If
End Sub

Private Sub chkCanal4_Click()
   If chkCanal4.Value = 0 Then
      nPrecioCanal4 = Round(nPrecioVenta * (1 + (nCanal4 / 100)), 2)
      txtCanal4.Text = Format(nPrecioCanal4, "###,##0." & Mid("0000000000", 1, nDecimal))
      txtCanal4.Enabled = False
   Else
      txtCanal4.Enabled = True
   End If
End Sub

Private Sub chkCanal5_Click()
   If chkCanal5.Value = 0 Then
      nPrecioCanal5 = Round(nPrecioVenta * (1 + (nCanal5 / 100)), 2)
      txtCanal5.Text = Format(nPrecioCanal5, "###,##0." & Mid("0000000000", 1, nDecimal))
      txtCanal5.Enabled = False
   Else
      txtCanal5.Enabled = True
   End If
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      cmdBoton(val(txtBoton)).Caption = Trim(txtBoton.Caption)
   End If
   cmdBoton(Index).backColor = &H80&
   cmdBoton(Index).Enabled = False
   cmdBoton(Index).Caption = txtResumido.Text
   txtBoton.Caption = Index
End Sub

Private Sub cmdNA_Click()
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      cmdBoton(val(txtBoton)).Caption = Trim(txtBoton.Caption)
      txtBoton.Caption = "NA"
   End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmProducto.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmProducto.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmProducto.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmProducto.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmProducto.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmProducto.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmProducto.RsCabecera.AbsolutePosition & " de " & frmProducto.RsCabecera.RecordCount
End Sub
Private Sub cmdOpcion_Click(Index As Integer)
frmProductoDetalle.SAPVALOR = False ' -- solo para opcion sap
Dim xGrupo As String
Dim xRGrupoAcceso As Recordset
xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
   If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
                '10305010 AGREGAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(0).Enabled = False
                End If
   End If
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
                cmdBuscaTipo.Enabled = True
               Blanquear Me
               chkActivo.Value = 1
               chkLCombinacion.Value = 0
               Iniciar
                      
                cmdAgregarFoto.Caption = "Editar"
                imgFoto.Picture = Nothing
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
          
          'GCAA 15032021
          
          
          Dim SaberResumido As Integer
          Dim SaberResumidoProducto As Integer
          
          SaberResumidoProducto = Calcular("select count(*)  as codigo from TPRODUCTO where tResumido ='" & txtResumido.Text & "' and tCodigoProducto <> '" & txtCodigo.Text & "'", Cn)
          SaberResumido = Calcular("select count(*)  as codigo from TPRODUCTO where tResumido ='" & txtResumido.Text & "'", Cn)
          
          If SaberResumidoProducto > 0 Then
                MsgBox "El Nombre Resumido ya se encuentra registrado", vbExclamation, sMensaje
                Exit Sub
          End If
          
          Select Case pais ' ok
            Case "000" 'Peru
            
            If txtPrecioBase.Caption = 0 Or txtPrecioBaseDelivery.Caption = 0 Or txtPrecioBaseLlevar.Caption = 0 Or txtPrecioBaseCanal4.Caption = 0 Or txtPrecioBaseCanal5.Caption = 0 Then
            
            Else
                Dim valObserva As Double
                valObserva = val(txtImpuesto1.Caption) + val(txtImpuesto1Delivery.Caption) + val(txtImpuesto1Llevar.Caption) + val(txtImpuesto1Canal4.Caption) + val(txtImpuesto1Canal5.Caption)
                
                If valObserva <= 0 Then
                  If ChtAplicaExoneracion.Value = 0 And ChtAplicaInafectacion.Value = 0 Then
                       MsgBox "El Producto no cuenta con IGV, debe seleccionar si es Exonerado o Inafecto", vbExclamation, sMensaje
                       Exit Sub
                  End If
                End If
                
             End If
                
            
            End Select
          
         
          
          
               Dim nCorrela As String
               Dim nPos As Variant
               
                
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If cboGrupo.Text = "" Then MsgBox "Ingrese el Grupo", vbExclamation, sMensaje: cboGrupo.SetFocus: Exit Sub
               If cboSubGrupo.Text = "" Then MsgBox "Ingrese el Sub Grupo", vbExclamation, sMensaje: cboSubGrupo.SetFocus: Exit Sub
               If sTipoProducto = "" Then MsgBox "Ingrese el Tipo de Producto", vbExclamation, sMensaje: cmdBuscaTipo.SetFocus: Exit Sub
               If cboArea.Text = "" Then MsgBox "Ingrese el área", vbExclamation, sMensaje: cboArea.SetFocus: Exit Sub
               If cboMoneda.Text = "" Then MsgBox "Ingrese la moneda", vbExclamation, sMensaje: cboMoneda.SetFocus: Exit Sub
               If cboUnidadNegocio.Text = "" Then MsgBox "Ingrese la unidad de negocio", vbExclamation, sMensaje: cboUnidadNegocio.SetFocus: Exit Sub
               If nPrecioVenta <= 0 Then MsgBox "El precio no puede ser cero", vbExclamation, sMensaje: txtPrecioVenta.SetFocus: Exit Sub
               If val(txtCantidadItemsMinima.Text) > val(txtCantidadItems.Text) Then MsgBox "La cantidad de items combo minima, no puede ser mayor a la cantidad de items combo maxima!!", vbInformation: txtCantidadItemsMinima.Text = txtCantidadItems.Text: Exit Sub
               If lSAP Then
                    If txtBarra.Text = "" Then MsgBox "Ingrese el Codigo Sap", vbExclamation, sMensaje: txtBarra.SetFocus: Exit Sub
               End If
                    
                    
                
                    
                    
                'insumocritico23
                If Me.chkInsumoCritico.Value = 1 Then
                
                    If SAPVALOR Then
                            scodigoInsumoCritico = LLenarInsumoCritico
                    Else
                        If Me.lblInsumo.Caption = "" Then
                            MsgBox "Ingrese el Insumo a enlazar con el plato de venta", vbExclamation, sMensaje
                            Exit Sub
                        End If
                    End If
                    
'                    If Me.lblInsumo.Caption = "" Then
'                            MsgBox "Ingrese el Insumo a enlazar con el plato de venta", vbExclamation, sMensaje
'                            Exit Sub
'
'                    End If
                    
                End If
                
               'Fin Modificacion
                If lAlmacen = True Then
                    If cboDescargo.BoundText = "D" Then
                         nExiste = Calcular("Select Count(*) As Codigo from vProducto Where lDDirecto=1 And tCodigoProducto = '" & txtDirecto.Text & "'", CnAlmacen)
                         If nExiste = 0 Then
                             MsgBox "El Descargo Directo es Incorrecto, Verifique!!!", vbExclamation, sMensaje
                             txtDirecto.SetFocus
                             Exit Sub
                         End If
                    End If
                    If cboDescargo.BoundText = "R" Then
                         nExiste = Calcular("Select Count(*) As Codigo from vRecetaVenta Where lActivo = 1 And tLocal='" & sLocal & "' And tRecetaVenta = '" & txtReceta.Text & "'", CnAlmacen)
                         If nExiste = 0 Then
                             MsgBox "El Descargo Receta es Incorrecto, Verifique!!!", vbExclamation, sMensaje
                             txtReceta.SetFocus
                             Exit Sub
                         End If
                    End If
               End If
               
               If chkLCombinacion.Value = 1 Then
                        If cboDescargo.BoundText = "D" Or cboDescargo.BoundText = "R" Then
                            If MsgBox("Esta indicando que el Producto es un Combo con Tipo de Descargo: " & UCase(cboDescargo.Text) & vbCrLf & _
                            " El Tipo de Descargo para un Combo es la Opción MENU, ¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                                Exit Sub
                            End If
                        End If
               End If
               
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoProducto) as Codigo from TPRODUCTO", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0000001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 7)
                  End If
                                    
                  If txtBoton.Caption <> "NA" Then
                     cmdBoton(val(txtBoton.Caption)).Caption = txtResumido.Text
                  End If
                   
                  If lInfhotel Then
                     Dim sTipoInf As String
                     sTipoInf = Calcular("select tValor as Codigo from vTipoProducto where Codigo='" & sTipoProducto & "' and lActivo=1", Cn)
                     If sTipoInf = "" Or sTipoInf = "0" Then
                        MsgBox "Error Fatal: Tipo de producto no asignado", vbCritical, sMensaje
                        Exit Sub
                     End If
                     ProductoInfhotel
                  Else
                     sCodInfhotel = ""
                  End If
                   
                   
                'AUDITORIA
                      
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TPRODUCTO", "PRODUCTOS", "01", sUsuario, sPasa, "", _
                   "tcodigoproducto", "CODIGO Producto", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tgrupo", "Grupo", cboGrupo.BoundText, "tsubgrupo", "Sub Grupo", cboSubGrupo.BoundText, "tarea", "Area Produccion", cboArea.BoundText, "tTipoProducto", "Tipo Producto", sTipoProducto, "tUnidadNegocio", "Unidad Negocio", cboUnidadNegocio.BoundText, "lImprimeArea", "Flag Imprime en Area", IIf(chkImprime.Value, "Verdadero", "Falso"), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tUsuario", "Usuario", sUsuario, "tBarra", "Codigo Barra", txtBarra.Text, "tAlternativa", "Descripcion Alternativa", txtAlternativa.Text, "ntiempo", "Tiempo Ideal Preparacion", val(txtTiempo.Text), _
                   "lCombinacion", "Flag Combinacion", IIf(chkLCombinacion.Value, "Verdadero", "Falso"), "nCombinacion", "Combinacion Items", val(txtCantidadItems.Text), "lMultiArea", "Flag MultiArea", IIf(chkMultiArea.Value, "Verdadero", "Falso"), "lPropiedad", "Flag Propiedad Visible", IIf(chkPropiedad.Value, "Verdadero", "Falso"), "lModificable", "Flag Precio Modificable", IIf(chkModificable.Value, "Verdadero", "Falso"), "lDescuento", "Permite Descuento", IIf(chkDescuento.Value, "Verdadero", "Falso"), "lBalanza", "Flag Ingreso Peso Con Balanza", IIf(chkBalanza.Value, "Verdadero", "Falso"), "nBOTON", "Botonera", val(txtBoton.Caption), "tMoneda", "Moneda", cboMoneda.BoundText, _
                   "limpuesto1", "Flag Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "lImpuesto2", "Flag Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "lImpuesto3", "Flag Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), "lImpuesto4", "Flag Impuesto 1 Canal 2", IIf(chkImpuesto1Delivery.Value, "Verdadero", "Falso"), "lImpuesto5", "Flag Impuesto 2 Canal 2", IIf(chkImpuesto2Delivery.Value, "Verdadero", "Falso"), "lImpuesto6", "Flag Impuesto 3 Canal 2", IIf(chkImpuesto3Delivery.Value, "Verdadero", "Falso"), "lImpuesto7", "Flag Impuesto 1 Canal 3", IIf(chkImpuesto1Llevar.Value, "Verdadero", "Falso"), "lImpuesto8", "Flag Impuesto 2 Canal 3", IIf(chkImpuesto2Llevar.Value, "Verdadero", "Falso"), "lImpuesto9", "Flag Impuesto 3 Canal 3", IIf(chkImpuesto3Llevar.Value, "Verdadero", "Falso"), "lImpuesto10", "Flag Impuesto 1 Canal 4", IIf(chkImpuesto1Canal4.Value, "Verdadero", "Falso"), "lImpuesto11", "Flag Impuesto 2 Canal 4", IIf(chkImpuesto2Canal4.Value, "Verdadero", "Falso"), _
                   "lImpuesto12", "Flag Impuesto 3 Canal 4", IIf(chkImpuesto3Canal4.Value, "Verdadero", "Falso"), "lImpuesto13", "Flag Impuesto 1 Canal 5", IIf(chkImpuesto1Canal5.Value, "Verdadero", "Falso"), "lImpuesto14", "Flag Impuesto 2 Canal 5", IIf(chkImpuesto2Canal5.Value, "Verdadero", "Falso"), "lImpuesto15", "Flag Impuesto 3 Canal 5", IIf(chkImpuesto3Canal5.Value, "Verdadero", "Falso"), "tDescargo", "Tipo Descargo", cboDescargo.BoundText, "nPrecioVenta", "Precio Venta Canal 1", val(nPrecioVenta), "npreciodelivery", "Precio Venta Canal 2", IIf(chkDelivery.Value, nPrecioDelivery, 0), "nPrecioLLevar", "Precio Venta Canal 3", IIf(chkLlevar.Value, nPrecioLlevar, 0), "nPreciocanal4", "Precio Venta Canal 4", IIf(chkCanal4.Value, nPrecioCanal4, 0), "nPrecioCanal5", "Precio Venta Canal 5", IIf(chkCanal5.Value, nPrecioCanal5, 0), "tEnlace", "Enlace Descargo", IIf(lAlmacen, IIf(cboDescargo.BoundText = "D", txtDirecto.Text, txtReceta.Text), txtEnlace.Text), _
                   "nInsumo", "Insumo Canal 1", nInsumo, "nInsumo2", "Insumo Canal 2", nInsumo2, "nInsumo3", "Insumo Canal 3", nInsumo3, "nInsumo4", "Insumo Canal 4", nInsumo4, "nInsumo5", "Insumo Canal 5", nInsumo5, "nGasto", "Gasto Canal 1", nGasto, "nGasto2", "Gasto Canal 2", nGasto2, "nGasto3", "Gasto Canal 3", nGasto3, "nGasto4", "Gasto Canal 4", nGasto4, "nGasto5", "Gasto Canal 5", nGasto5, "nManoObra", "Gasto Mano Obra Canal 1", nMObra, "nManoObra2", "Gasto Mano Obra Canal 2", nMObra2, "nManoObra3", "Gasto Mano Obra Canal 3", nMObra3, "nmanoobra4", "Gasto Mano Obra Canal 4", nMObra4, "nManoObra5", "Gasto Mano Obra Canal 5", nMObra5, "tInfhotel", "Codigo Infhotel", sCodInfhotel, _
                   "lLocal", "Visualiza en Canal 1", IIf(chkVisualizarLocal.Value, "Verdadero", "Falso"), "lDelivery", "Visualiza en Canal 2", IIf(chkVisualizarDelivery.Value, "Verdadero", "Falso"), "lLlevar", "Visualiza en Canal 3", IIf(chkVisualizarLlevar.Value, "Verdadero", "Falso"), "lCanal4", "Visualiza en Canal 4", IIf(chkvisualizarCanal4.Value, "Verdadero", "Falso"), "lCanal5", "Visualiza en Canal 5", IIf(chkVisualizarCanal5.Value, "Verdadero", "Falso"), "lControlInsumoCritico", "Flag Control Insumo Critico", IIf(chkInsumoCritico.Value, "Verdadero", "Falso"), "tCodigoInsumo", "Codigo Insumo Critico", IIf(chkInsumoCritico.Value = 1, scodigoInsumoCritico, ""), "lDefineAreaProduccion", "Flag Define Area Produccion", IIf(Me.chkDefineArea.Value, "Verdadero", "Falso"), "tCodigoProductoSunat", "CodigoProductoSunat", Me.lblProductoSunat.Caption, "nPrecioVentaRefer", "Precio de Venta Referencial", Me.txtPrecioReferencial.Text, _
                   "nCombinacionMinima", "Cantidad de Items Combos Minimo", Me.txtCantidadItemsMinima.Text, "tCodigoSKU", "Codigo Externo SKU", Me.txtSKU.Text, "lLey1", "Impuesto ley 1", IIf(Me.chkLey1.Value, "Verdadero", "Falso"), "lActivahomologacionexterno", "Uso Externo por web u Otro", IIf(Me.chkActivaUsoExterno.Value, "Verdadero", "Falso"), "tdescripcionpatio", "Descripcion para patio de comidas", Trim(Me.txtDesPatio.Text), "lActivaPatioComida", "Aciva check para patio de comidas", IIf(Me.chkActivaPatioComida.Value, "Verdadero", "Falso"), "tNombrePatio", "Nombre que viaja a patio de comidas", Trim(Me.txtNombrePatio.Text), "lComboVisible", "Combo visible", IIf(chklComboVisible.Value, "Verdadero", "Falso"), "IntegracionWebApi", "Integracion WebApi", IIf(chkIntegracionWebApi.Value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                   'AUDITORIA
                   
                  'Cambiar el SQL
                  Isql = "insert into TPRODUCTO( " & _
                         "tCodigoProducto, tDetallado, tResumido, tGrupo, tSubGrupo, tUnidadNegocio, tTipoProducto, " & _
                         "lImpuesto1, lImpuesto2, lImpuesto3, lImpuesto4, lImpuesto5, lImpuesto6, lImpuesto7, lImpuesto8, lImpuesto9, lImpuesto10, lImpuesto11, lImpuesto12, lImpuesto13, lImpuesto14, lImpuesto15, " & _
                         "nPrecioVenta, nPrecioDelivery, nPrecioLlevar, nPrecioCanal4, nPrecioCanal5, nInsumo, nGasto, nManoObra, nInsumo2, nGasto2, nManoObra2, nInsumo3, nGasto3, nManoObra3, nInsumo4, nGasto4, nManoObra4, nInsumo5, nGasto5, nManoObra5, " & _
                         "nBoton, lCombinacion, nCombinacion, tArea, lImprimeArea, lModificable, tDescargo, tEnlace, tMoneda, tBarra, lPropiedad, lActivo, tInfhotel, lDescuento, lLocal, lDelivery, lLlevar, lCanal4, lCanal5, tUsuario, fRegistro,lmultiarea, " & _
                         "talternativa,lControlInsumoCritico,tcodigoInsumo, nTiempo, lBalanza, lSolicitaAutorizacionEnvio,lDefineAreaProduccion, tCodigoProductoSunat, nPrecioVentaRefer, nCombinacionMinima, tCodigoSKU, lley1, lActivahomologacionExterno,surlimage,tAplicaExoneracion,  tAplicaInafectacion, tNombrePatio,lComboVisible, IntegracionWebApi) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', '" & txtResumido.Text & "', " & _
                                " '" & cboGrupo.BoundText & "', '" & cboSubGrupo.BoundText & "', " & _
                                " '" & cboUnidadNegocio.BoundText & "', '" & sTipoProducto & "', " & _
                                       chkImpuesto1.Value & ", " & chkImpuesto2.Value & ", " & chkImpuesto3.Value & ", " & _
                                       chkImpuesto1Delivery.Value & ", " & chkImpuesto2Delivery.Value & ", " & chkImpuesto3Delivery.Value & ", " & chkImpuesto1Llevar.Value & ", " & chkImpuesto2Llevar.Value & ", " & chkImpuesto3Llevar.Value & ", " & _
                                       chkImpuesto1Canal4.Value & ", " & chkImpuesto2Canal4.Value & ", " & chkImpuesto3Canal4.Value & ", " & chkImpuesto1Canal5.Value & ", " & chkImpuesto2Canal5.Value & ", " & chkImpuesto3Canal5.Value & ", " & _
                                       nPrecioVenta & ", " & IIf(chkDelivery.Value = 1, nPrecioDelivery, 0) & ", " & IIf(chkLlevar.Value = 1, nPrecioLlevar, 0) & ", " & IIf(chkCanal4.Value = 1, nPrecioCanal4, 0) & ", " & IIf(chkCanal5.Value = 1, nPrecioCanal5, 0) & ", " & _
                                       nInsumo & ", " & nGasto & ", " & nMObra & ", " & nInsumo2 & ", " & nGasto2 & ", " & nMObra2 & ", " & nInsumo3 & ", " & nGasto3 & ", " & nMObra3 & ", " & nInsumo4 & ", " & nGasto4 & ", " & nMObra4 & ", " & nInsumo5 & ", " & nGasto5 & ", " & nMObra5 & ", " & _
                                       val(txtBoton.Caption) & ", " & chkLCombinacion.Value & ", " & val(txtCantidadItems.Text) & ", " & _
                                " '" & cboArea.BoundText & "', " & _
                                       chkImprime.Value & ", " & chkModificable.Value & ", " & _
                                " '" & cboDescargo.BoundText & "', " & _
                                " '" & IIf(lAlmacen, IIf(cboDescargo.BoundText = "D", txtDirecto.Text, txtReceta.Text), txtEnlace.Text) & "', " & _
                                " '" & cboMoneda.BoundText & "', " & _
                                " '" & txtBarra.Text & "', " & chkPropiedad & ", " & chkActivo.Value & ", '" & sCodInfhotel & "', " & chkDescuento.Value & ", " & _
                                       chkVisualizarLocal.Value & ", " & chkVisualizarDelivery.Value & ", " & chkVisualizarLlevar.Value & ", " & chkvisualizarCanal4.Value & ", " & chkVisualizarCanal5.Value & ", " & _
                                " '" & sUsuario & "', getdate(), " & chkMultiArea.Value & ",'" & Trim(Me.txtAlternativa.Text) & "'," & chkInsumoCritico.Value & ",'" & IIf(chkInsumoCritico.Value = 1, scodigoInsumoCritico, "") & "', " & CInt(txtTiempo.Text) & _
                                ", " & chkBalanza.Value & ", " & chkSolicitaAutorizacionEnvio.Value & " , " & chkDefineArea.Value & ",'" & Me.lblProductoSunat.Caption & "'," & CDbl(Me.txtPrecioReferencial.Text) & ", " & val(Me.txtCantidadItemsMinima.Text) & " , '" & Me.txtSKU.Text & "', " & Me.chkLey1.Value & ", " & Me.chkActivaUsoExterno.Value & ", " & " '" & txturlimage.Text & "', " & Me.ChtAplicaExoneracion.Value & ", " & Me.ChtAplicaInafectacion.Value & ",'" & Trim(Me.txtNombrePatio.Text) & "', " & Me.chklComboVisible.Value & ", " & Me.chkIntegracionWebApi.Value & ")"
                  Cn.Execute Isql
                  
                  Cn.Execute "update tproducto set lNuevoModificado=1, lActivaPatioComida=" & Me.chkActivaPatioComida.Value & " , tDescripcionPatio='" & Trim(txtDesPatio.Text) & "', tStockPlatoCritico='" & IIf(txtStockPlatoCritico.Text = "", 0, txtStockPlatoCritico.Text) & "' where tcodigoproducto='" & txtCodigo & "'"
                  
                  If chkImprime.Value Then
                     Cn.Execute "insert into TPRODUCTOAREA (tCodigoProducto, tArea) values ('" & txtCodigo.Text & "', '" & cboArea.BoundText & "')"
                     RsGImpresion.Requery
                     RsGImpresion.Filter = "tCodigoProducto='" & txtCodigo.Text & "'"
                  End If
                                    
                  RsBoton.Requery
                  frmProducto.RsCabecera.Sort = "Codigo ASC"
                  frmProducto.RsCabecera.Requery
                  Sw = False
                  frmProducto.RsCabecera.MoveLast
                  Asignar
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmProducto.RsCabecera.RecordCount = 0, 0, frmProducto.RsCabecera.AbsolutePosition) & " de " & frmProducto.RsCabecera.RecordCount
                  
                  If SAPVALOR Then
                    'Call LLenarInsumoCritico
                    Cn.Execute "exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sPasa & "','" & sCodSapInf & "','',4"
                  End If
                                    
               Else
                'AUDITORIA
                      
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TPRODUCTO", "PRODUCTOS", "02", sUsuario, sPasa, "", _
                   "tcodigoproducto", "CODIGO Producto", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tgrupo", "Grupo", cboGrupo.BoundText, "tsubgrupo", "Sub Grupo", cboSubGrupo.BoundText, "tarea", "Area Produccion", cboArea.BoundText, "tTipoProducto", "Tipo Producto", sTipoProducto, "tUnidadNegocio", "Unidad Negocio", cboUnidadNegocio.BoundText, "lImprimeArea", "Flag Imprime en Area", IIf(chkImprime.Value, "Verdadero", "Falso"), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tUsuario", "Usuario", sUsuario, "tBarra", "Codigo Barra", txtBarra.Text, "tAlternativa", "Descripcion Alternativa", txtAlternativa.Text, "ntiempo", "Tiempo Ideal Preparacion", val(txtTiempo.Text), _
                   "lCombinacion", "Flag Combinacion", IIf(chkLCombinacion.Value, "Verdadero", "Falso"), "nCombinacion", "Combinacion Items", val(txtCantidadItems.Text), "lMultiArea", "Flag MultiArea", IIf(chkMultiArea.Value, "Verdadero", "Falso"), "lPropiedad", "Flag Propiedad Visible", IIf(chkPropiedad.Value, "Verdadero", "Falso"), "lModificable", "Flag Precio Modificable", IIf(chkModificable.Value, "Verdadero", "Falso"), "lDescuento", "Permite Descuento", IIf(chkDescuento.Value, "Verdadero", "Falso"), "lBalanza", "Flag Ingreso Peso Con Balanza", IIf(chkBalanza.Value, "Verdadero", "Falso"), "nBOTON", "Botonera", val(txtBoton.Caption), "tMoneda", "Moneda", cboMoneda.BoundText, _
                   "limpuesto1", "Flag Impuesto 1", IIf(chkImpuesto1.Value, "Verdadero", "Falso"), "lImpuesto2", "Flag Impuesto 2", IIf(chkImpuesto2.Value, "Verdadero", "Falso"), "lImpuesto3", "Flag Impuesto 3", IIf(chkImpuesto3.Value, "Verdadero", "Falso"), "lImpuesto4", "Flag Impuesto 1 Canal 2", IIf(chkImpuesto1Delivery.Value, "Verdadero", "Falso"), "lImpuesto5", "Flag Impuesto 2 Canal 2", IIf(chkImpuesto2Delivery.Value, "Verdadero", "Falso"), "lImpuesto6", "Flag Impuesto 3 Canal 2", IIf(chkImpuesto3Delivery.Value, "Verdadero", "Falso"), "lImpuesto7", "Flag Impuesto 1 Canal 3", IIf(chkImpuesto1Llevar.Value, "Verdadero", "Falso"), "lImpuesto8", "Flag Impuesto 2 Canal 3", IIf(chkImpuesto2Llevar.Value, "Verdadero", "Falso"), "lImpuesto9", "Flag Impuesto 3 Canal 3", IIf(chkImpuesto3Llevar.Value, "Verdadero", "Falso"), "lImpuesto10", "Flag Impuesto 1 Canal 4", IIf(chkImpuesto1Canal4.Value, "Verdadero", "Falso"), "lImpuesto11", "Flag Impuesto 2 Canal 4", IIf(chkImpuesto2Canal4.Value, "Verdadero", "Falso"), _
                   "lImpuesto12", "Flag Impuesto 3 Canal 4", IIf(chkImpuesto3Canal4.Value, "Verdadero", "Falso"), "lImpuesto13", "Flag Impuesto 1 Canal 5", IIf(chkImpuesto1Canal5.Value, "Verdadero", "Falso"), "lImpuesto14", "Flag Impuesto 2 Canal 5", IIf(chkImpuesto2Canal5.Value, "Verdadero", "Falso"), "lImpuesto15", "Flag Impuesto 3 Canal 5", IIf(chkImpuesto3Canal5.Value, "Verdadero", "Falso"), "tDescargo", "Tipo Descargo", cboDescargo.BoundText, "nPrecioVenta", "Precio Venta Canal 1", val(nPrecioVenta), "npreciodelivery", "Precio Venta Canal 2", IIf(chkDelivery.Value, nPrecioDelivery, 0), "nPrecioLLevar", "Precio Venta Canal 3", IIf(chkLlevar.Value, nPrecioLlevar, 0), "nPreciocanal4", "Precio Venta Canal 4", IIf(chkCanal4.Value, nPrecioCanal4, 0), "nPrecioCanal5", "Precio Venta Canal 5", IIf(chkCanal5.Value, nPrecioCanal5, 0), "tEnlace", "Enlace Descargo", IIf(lAlmacen, IIf(cboDescargo.BoundText = "D", txtDirecto.Text, txtReceta.Text), txtEnlace.Text), _
                   "nInsumo", "Insumo Canal 1", nInsumo, "nInsumo2", "Insumo Canal 2", nInsumo2, "nInsumo3", "Insumo Canal 3", nInsumo3, "nInsumo4", "Insumo Canal 4", nInsumo4, "nInsumo5", "Insumo Canal 5", nInsumo5, "nGasto", "Gasto Canal 1", nGasto, "nGasto2", "Gasto Canal 2", nGasto2, "nGasto3", "Gasto Canal 3", nGasto3, "nGasto4", "Gasto Canal 4", nGasto4, "nGasto5", "Gasto Canal 5", nGasto5, "nManoObra", "Gasto Mano Obra Canal 1", nMObra, "nManoObra2", "Gasto Mano Obra Canal 2", nMObra2, "nManoObra3", "Gasto Mano Obra Canal 3", nMObra3, "nmanoobra4", "Gasto Mano Obra Canal 4", nMObra4, "nManoObra5", "Gasto Mano Obra Canal 5", nMObra5, "tInfhotel", "Codigo Infhotel", sCodInfhotel, _
                   "lLocal", "Visualiza en Canal 1", IIf(chkVisualizarLocal.Value, "Verdadero", "Falso"), "lDelivery", "Visualiza en Canal 2", IIf(chkVisualizarDelivery.Value, "Verdadero", "Falso"), "lLlevar", "Visualiza en Canal 3", IIf(chkVisualizarLlevar.Value, "Verdadero", "Falso"), "lCanal4", "Visualiza en Canal 4", IIf(chkvisualizarCanal4.Value, "Verdadero", "Falso"), "lCanal5", "Visualiza en Canal 5", IIf(chkVisualizarCanal5.Value, "Verdadero", "Falso"), "lControlInsumoCritico", "Flag Control Insumo Critico", IIf(chkInsumoCritico.Value, "Verdadero", "Falso"), "tCodigoInsumo", "Codigo Insumo Critico", IIf(chkInsumoCritico.Value = 1, scodigoInsumoCritico, ""), "lDefineAreaProduccion", "Flag Define Area Produccion", IIf(Me.chkDefineArea.Value, "Verdadero", "Falso"), "tCodigoProductoSunat", "CodigoProductoSunat", Me.lblProductoSunat.Caption, "nPrecioVentaRefer", "Precio de Venta Referencial", Me.txtPrecioReferencial.Text, _
                   "nCombinacionMinima", "Cantidad de Items Combos Minimo", Me.txtCantidadItemsMinima.Text, "tCodigoSKU", "Codigo Externo SKU", Me.txtSKU.Text, "lLey1", "Impuesto ley 1", IIf(Me.chkLey1.Value, "Verdadero", "Falso"), "lActivahomologacionexterno", "Uso Externo por web u Otro", IIf(Me.chkActivaUsoExterno.Value, "Verdadero", "Falso"), "tdescripcionpatio", "Descripcion para patio de comidas", Trim(Me.txtDesPatio.Text), "lActivaPatioComida", "Aciva check para patio de comidas", IIf(Me.chkActivaPatioComida.Value, "Verdadero", "Falso"), "tNombrePatio", "Nombre que viaja a patio de comidas", Trim(Me.txtNombrePatio.Text), "lComboVisible", "Combo visible", IIf(chklComboVisible.Value, "Verdadero", "Falso"), "IntegracionWebApi", "Activa Integracion Web Api", IIf(Me.chkIntegracionWebApi.Value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                Dim xRecetaAntes As String
                Dim xTipoDescargo As String
                Dim countReceta As Integer
                If lAlmacen Then
                    xRecetaAntes = Calcular("Select tEnlace As codigo From TPRODUCTO where tCodigoProducto='" & txtCodigo.Text & "'", Cn)
                    xTipoDescargo = Calcular("Select tDescargo As codigo From TPRODUCTO where tCodigoProducto='" & txtCodigo.Text & "'", Cn)
                    countReceta = Calcular("Select count(*) as codigo From TPRODUCTO where tEnlace='" & xRecetaAntes & "'", Cn)
                    
                    If xTipoDescargo = "R" And cboDescargo.BoundText <> xTipoDescargo Then
                        If countReceta > 1 Then
                            MsgBox "La receta " & xRecetaAntes & " no se desactivará debido a que se encuentra enlazada a otro producto de venta", vbInformation, sMensaje
                        Else
                            Isql = "Update MRECETAVENTA Set lActivo = 0 Where tRecetaVenta = '" & xRecetaAntes & "' and tLocal = '" & sLocal & "'"
                            CnAlmacen.Execute Isql
                        End If

                    End If
                End If
               
                  'Cambiar el SQL
                   Isql = "update TPRODUCTO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', tResumido ='" & txtResumido.Text & "', " & _
                         "talternativa ='" & txtAlternativa.Text & "', " & _
                         "tGrupo ='" & cboGrupo.BoundText & "', " & _
                         "tSubGrupo ='" & cboSubGrupo.BoundText & "', tUnidadNegocio ='" & cboUnidadNegocio.BoundText & "', " & _
                         "tTipoProducto ='" & sTipoProducto & "', " & _
                         "tArea ='" & cboArea.BoundText & "',  lImprimeArea =" & chkImprime.Value & ", " & _
                         "lModificable =" & chkModificable.Value & ", " & _
                         "lBalanza =" & chkBalanza.Value & ", " & _
                         "lImpuesto1 = " & chkImpuesto1.Value & ", lImpuesto2 = " & chkImpuesto2.Value & ", lImpuesto3 = " & chkImpuesto3.Value & ", lImpuesto4 = " & chkImpuesto1Delivery.Value & ", lImpuesto5 = " & chkImpuesto2Delivery.Value & ", lImpuesto6 = " & chkImpuesto3Delivery.Value & ", lImpuesto7 = " & chkImpuesto1Llevar.Value & ", lImpuesto8 = " & chkImpuesto2Llevar.Value & ", lImpuesto9 = " & chkImpuesto3Llevar.Value & ", " & _
                         "lImpuesto10 = " & chkImpuesto1Canal4.Value & ", lImpuesto11 = " & chkImpuesto2Canal4.Value & ", lImpuesto12 = " & chkImpuesto3Canal4.Value & ", lImpuesto13 = " & chkImpuesto1Canal5.Value & ", lImpuesto14 = " & chkImpuesto2Canal5.Value & ", lImpuesto15 = " & chkImpuesto3Canal5.Value & ", " & _
                         "nPrecioVenta = " & nPrecioVenta & ", nPrecioDelivery = " & IIf(chkDelivery.Value = 1, nPrecioDelivery, 0) & ", nPrecioLlevar = " & IIf(chkLlevar.Value = 1, nPrecioLlevar, 0) & ", nPrecioCanal4 = " & IIf(chkCanal4.Value = 1, nPrecioCanal4, 0) & ", nPreciocanal5 = " & IIf(chkCanal5.Value = 1, nPrecioCanal5, 0) & ", " & _
                         "nInsumo = " & nInsumo & ", nGasto = " & nGasto & ", nManoObra = " & nMObra & ", " & _
                         "nInsumo2 = " & nInsumo2 & ", nGasto2 = " & nGasto2 & ", nManoObra2 = " & nMObra2 & ", " & _
                         "nInsumo3 = " & nInsumo3 & ", nGasto3 = " & nGasto3 & ", nManoObra3 = " & nMObra3 & ", " & _
                         "nInsumo4 = " & nInsumo4 & ", nGasto4 = " & nGasto4 & ", nManoObra4 = " & nMObra4 & ", " & _
                         "nInsumo5 = " & nInsumo5 & ", nGasto5 = " & nGasto5 & ", nManoObra5 = " & nMObra5 & ", nTiempo=" & CInt(txtTiempo.Text) & ", " & _
                         "nBoton = " & val(txtBoton.Caption) & ", lCombinacion = " & chkLCombinacion.Value & ", nCombinacion = " & val(txtCantidadItems.Text) & ", " & _
                         "tMoneda = '" & cboMoneda.BoundText & "', tDescargo = '" & cboDescargo.BoundText & "', " & _
                         "tEnlace = '" & IIf(lAlmacen, IIf(cboDescargo.BoundText = "D", txtDirecto.Text, txtReceta.Text), txtEnlace.Text) & "', " & _
                         "tBarra = '" & txtBarra.Text & "', lPropiedad =" & chkPropiedad.Value & " , ldefineareaproduccion=" & Me.chkDefineArea.Value & ", " & _
                         "lActivo =" & chkActivo.Value & ", lDescuento =" & chkDescuento.Value & ",lmultiarea =" & chkMultiArea.Value & ", " & _
                         "fRegistro=getdate(), tUsuario='" & sUsuario & "', lSolicitaAutorizacionEnvio = " & chkSolicitaAutorizacionEnvio.Value & ", " & _
                         "lLocal =" & chkVisualizarLocal.Value & ", lDelivery =" & chkVisualizarDelivery.Value & ", lLlevar =" & chkVisualizarLlevar.Value & ", lCanal4 =" & chkvisualizarCanal4.Value & ", lCanal5 =" & chkVisualizarCanal5.Value & _
                         ",lreplica=1,lControlInsumoCritico=" & chkInsumoCritico.Value & ", tcodigoInsumo='" & IIf(chkInsumoCritico.Value = 1, scodigoInsumoCritico, "") & "', tCodigoProductosunat='" & Me.lblProductoSunat.Caption & "', nPrecioVentaRefer= " & CDbl(Me.txtPrecioReferencial.Text) & " , nCombinacionMinima = " & val(txtCantidadItemsMinima.Text) & ", tCodigoSKU='" & Me.txtSKU.Text & "', lLey1=" & Me.chkLey1.Value & ", lActivaHomologacionexterno=" & Me.chkActivaUsoExterno.Value & " , surlimage='" & txturlimage.Text & "' , tAplicaExoneracion = " & Me.ChtAplicaExoneracion.Value & ", tAplicaInafectacion = " & Me.ChtAplicaInafectacion.Value & "  where tCodigoProducto = '" & txtCodigo & "' "
                  Cn.Execute Isql
                  
                  Cn.Execute "update tproducto set lNuevoModificado=1,lComboVisible=" & Me.chklComboVisible.Value & " , lActivaPatioComida=" & Me.chkActivaPatioComida.Value & " , tDescripcionPatio='" & Trim(txtDesPatio.Text) & "', tStockPlatoCritico='" & IIf(txtStockPlatoCritico.Text = "", 0, txtStockPlatoCritico.Text) & "', tnombrepatio='" & Trim(Me.txtNombrePatio.Text) & "', IntegracionWebApi='" & Trim(Me.chkIntegracionWebApi.Value) & "' where tcodigoproducto='" & txtCodigo & "'"
                  
                  nPos = frmProducto.RsCabecera.Bookmark
                  frmProducto.RsCabecera.Requery
                  RsBoton.Requery
                  RsGImpresion.Requery
                  If frmProducto.RsCabecera.RecordCount = 0 Then
                     frmProducto.RsCabecera.Filter = adFilterNone
                  End If
                  frmProducto.RsCabecera.Bookmark = nPos
                  
                  If chkLCombinacion.Value Then
                      Cn.Execute "delete from TPRODUCTOAREA where tCodigoProducto='" & txtCodigo.Text & "'"
                  End If
                  
                  If lInfhotel Then
                     ProductoInfhotel
                  End If
                                    
                  'Graba Oferta
                  If RsOferta.RecordCount > 0 Then
                     nPos = RsOferta.Bookmark
                     Cn.Execute "update TOFERTA " & _
                                "set  tHoraInicial='" & Format(dtpHoraIni.Value, "HH:nn") & "'," & _
                                " tHoraFinal='" & Format(dtpHoraFin.Value, "HH:nn") & "'," & _
                                " nRatio='" & txtRatio.Text & "'," & _
                                " lAcumulable='" & chkAcumulable.Value & "'," & _
                                " lActivo='" & ChkActv.Value & "'" & _
                                "where tcodigoproducto = '" & Trim(txtCodigo.Text) & "' and tOferta='" & grdOferta.Columns(0).Value & "'"
                     RsOferta.Requery
                     RsOferta.Bookmark = nPos
                  End If
                  Asignar
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
               '********** Para Administracion Centralizada / Patio Comida **************
               'Cn.Execute "update tproducto set lNuevoModificado=1, lActivaPatioComida=" & Me.chkActivaPatioComida.Value & "  where tcodigoproducto='" & txtCodigo & "'"
               
               cmdTexto.Caption = "Registro " & IIf(frmProducto.RsCabecera.RecordCount = 0, 0, frmProducto.RsCabecera.AbsolutePosition) & " de " & frmProducto.RsCabecera.RecordCount
          
          Case Is = 2 ' Eliminar
               If frmProducto.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Producto " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select codigo as Codigo From vproducto where lactivo=1 AND  codigo=(SELECT tCodItemCover FROM TPARAMETRO)", Cn) > 0 Then
                  MsgBox "Este Producto esta Enlazado como ITEM COVER en parametros generales, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If Calcular("select count(tCodigoProducto) as Codigo From DPEDIDO where tCodigoProducto='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este producto, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               
               If Calcular("select count(tcodigoproducto) as codigo from TCOMBO where tcodigoproducto='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Este producto esta registrado como Item de al menos un Combo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                   'AUDITORIA
                      
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TPRODUCTO", "PRODUCTOS", "03", sUsuario, sPasa, "", _
                   "tcodigoproducto", "CODIGO Producto", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                   'AUDITORIA
                                  
               
               
               'Cambia el Delete
               Cn.Execute "delete from TPRODUCTO where tCodigoProducto = '" & txtCodigo & "'"
               Cn.Execute "delete from TPRODUCTOAREA where tCodigoProducto = '" & txtCodigo & "'"
               Cn.Execute "delete from TCOMBO where tCombo='" & txtCodigo.Text & "'"
               frmProducto.RsCabecera.Requery
               If frmProducto.RsCabecera.RecordCount <> 0 Then
                  frmProducto.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmProducto.RsCabecera.RecordCount = 0, 0, frmProducto.RsCabecera.AbsolutePosition) & " de " & frmProducto.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
               
               'Delete infhotel
               If lInfhotel Then
                  'CnInfhotel.Execute "delete from TPRODUCTOPRECIO where tCodigo"
                  'CnInfhotel.Execute "delete from TPRODUCTOPRECIO where tCodigo"
               End If
         
          Case Is = 3 ' Salir
               Unload Me
               
          Case Is = 4 ' Salir
               frmProductoBoton.Show vbModal
                              
          Case Is = 5 ' Boton Recetas
               sCodigo = txtReceta.Text
               frmRecetaVentaDetalle.Show vbModal
          
   End Select

End Sub

Private Sub dtpHoraFin_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
   If dtpHoraIni.Value > dtpHoraFin.Value Then
      dtpHoraFin.Value = dtpHoraFin.Value
      MsgBox "Error en Horas", vbExclamation, sMensaje
   End If
End Sub

Private Sub dtpHoraIni_Change()
   If dtpHoraIni.Value > dtpHoraFin.Value Then
      dtpHoraIni.Value = dtpHoraFin.Value
      MsgBox "Error en Horas", vbExclamation, sMensaje
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'insumocritico23
    If Calcular("select isnull(lprinter,0) as codigo from tparametro  ", Cn) = False Then
        tabProducto.TabEnabled(5) = False
    End If
    'insumocritico23
    
    If lInfhotel Then
    
        If Sw = False Then
             cmdBuscaTipo.Enabled = False
        Else
             cmdBuscaTipo.Enabled = True
        End If
        
       Dim sInfhotelRuta As String
       Dim sInfhotelMDB As String
       Dim sCajaInfhotel As String
      
       sInfhotelRuta = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "SERVIDOR", "."))
       sInfhotelMDB = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "BASEDATO", "INFHOTEL"))
       sCajaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "CAJA", "01"))
       sPuntoVentaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "PUNTOVENTA", "01"))
       
       Set CnInfhotel = New Connection
       CnInfhotel.Provider = "SQLOLEDB"
       CnInfhotel.CursorLocation = adUseServer
       CnInfhotel.ConnectionString = "User ID=" & sUserName & _
                                     ";password=" & sUserPassword & _
                                     ";Data Source=" & sInfhotelRuta & _
                                     ";Initial Catalog=" & sInfhotelMDB
       CnInfhotel.Open
       sHotel = Calcular("select tHotel as Codigo from vCaja where tCaja='" & sCajaInfhotel & "'", CnInfhotel)
       sHotel = IIf(sHotel = "00", "01", sHotel)
    End If
    
    'Seguridad Me
    xCantidad = 0
        
    'Ingrese el SubTitulo
    Me.Caption = " Productos y Precios "
        
    TabCostos.TabCaption(0) = sBoton1
    TabCostos.Tab = 0
    
    lblLocal.Caption = " " & sBoton1 & " "
    
    If sBoton2 <> "" Then
       chkDelivery.Caption = " " & sBoton2 & " "
       TabCostos.TabCaption(1) = sBoton2
    Else
       Frame2.Enabled = False
       chkDelivery.Enabled = False
       TabCostos.TabEnabled(1) = False
    End If
    If sBoton3 <> "" Then
        chkLlevar.Caption = " " & sBoton3 & " "
        TabCostos.TabCaption(2) = sBoton3
    Else
       Frame3.Enabled = False
       chkLlevar.Enabled = False
       TabCostos.TabEnabled(2) = False
    End If
    If sBoton4 <> "" Then
       TabCostos.TabCaption(3) = sBoton4
       chkCanal4.Caption = " " & sBoton4 & " "
    Else
       Frame4.Enabled = False
       chkCanal4.Enabled = False
       TabCostos.TabEnabled(3) = False
    End If
    If sBoton5 <> "" Then
       chkCanal5.Caption = " " & sBoton5 & " "
       TabCostos.TabCaption(4) = sBoton5
    Else
       Frame5.Enabled = False
       chkCanal5.Enabled = False
       TabCostos.TabEnabled(4) = False
    End If
    
    chkVisualizarLocal.Caption = " Visualiza " & sBoton1 & " "
    chkVisualizarDelivery.Caption = " Visualiza " & sBoton2 & " "
    chkVisualizarLlevar.Caption = " Visualiza " & sBoton3 & " "
    chkvisualizarCanal4.Caption = " Visualiza " & sBoton4 & " "
    chkVisualizarCanal5.Caption = " Visualiza " & sBoton5 & " "
    
    If sImpuesto1 = "" Then
       lblI1.Visible = False
       txtImpuesto1.Visible = False
       chkImpuesto1.Visible = False
       chkImpuesto1.Value = 0
              
       lblI1Delivery.Visible = False
       txtImpuesto1Delivery.Visible = False
       chkImpuesto1Delivery.Visible = False
       chkImpuesto1Delivery.Value = 0
       
       lblI1Llevar.Visible = False
       txtImpuesto1Llevar.Visible = False
       chkImpuesto1Llevar.Visible = False
       chkImpuesto1Llevar.Value = 0
              
       lblI1Canal4.Visible = False
       txtImpuesto1Canal4.Visible = False
       chkImpuesto1Canal4.Visible = False
       chkImpuesto1Canal4.Value = 0
       
       lblI1Canal5.Visible = False
       txtImpuesto1Canal5.Visible = False
       chkImpuesto1Canal5.Visible = False
       chkImpuesto1Canal5.Value = 0
    Else
       lblI1.Caption = sImpuesto1 & " : "
       lblI1Llevar.Caption = sImpuesto1 & " : "
       lblI1Delivery.Caption = sImpuesto1 & " : "
       lblI1Canal4.Caption = sImpuesto1 & " : "
       lblI1Canal5.Caption = sImpuesto1 & " : "
    End If
    
    If sImpuesto2 = "" Then
       lblI2.Visible = False
       txtImpuesto2.Visible = False
       chkImpuesto2.Visible = False
       chkImpuesto2.Value = 0
       
       lblI2Llevar.Visible = False
       txtImpuesto2Llevar.Visible = False
       chkImpuesto2Llevar.Visible = False
       chkImpuesto2Llevar.Value = 0
       
       lblI2Delivery.Visible = False
       txtImpuesto2Delivery.Visible = False
       chkImpuesto2Delivery.Visible = False
       chkImpuesto2Delivery.Value = 0
       
       lblI2Canal4.Visible = False
       txtImpuesto2Canal4.Visible = False
       chkImpuesto2Canal4.Visible = False
       chkImpuesto2Canal4.Value = 0
       
       lblI2Canal5.Visible = False
       txtImpuesto2Canal5.Visible = False
       chkImpuesto2Canal5.Visible = False
       chkImpuesto2Canal5.Value = 0
    Else
       lblI2.Caption = sImpuesto2 & " : "
       lblI2Llevar.Caption = sImpuesto2 & " : "
       lblI2Delivery.Caption = sImpuesto2 & " : "
       lblI2Canal4.Caption = sImpuesto2 & " : "
       lblI2Canal5.Caption = sImpuesto2 & " : "
    End If
    
    If sImpuesto3 = "" Then
       lblI3.Visible = False
       txtImpuesto3.Visible = False
       chkImpuesto3.Visible = False
       chkImpuesto3.Value = 0
       
       lblI3Llevar.Visible = False
       txtImpuesto3Llevar.Visible = False
       chkImpuesto3Llevar.Visible = False
       chkImpuesto3Llevar.Value = 0
       
       lblI3Delivery.Visible = False
       txtImpuesto3Delivery.Visible = False
       chkImpuesto3Delivery.Visible = False
       chkImpuesto3Delivery.Value = 0
       
       lblI3Canal4.Visible = False
       txtImpuesto3Canal4.Visible = False
       chkImpuesto3Canal4.Visible = False
       chkImpuesto3Canal4.Value = 0
       
       lblI3Canal5.Visible = False
       txtImpuesto3Canal5.Visible = False
       chkImpuesto3Canal5.Visible = False
       chkImpuesto3Canal5.Value = 0
    Else
       lblI3.Caption = sImpuesto3 & " : "
       lblI3Llevar.Caption = sImpuesto3 & " : "
       lblI3Delivery.Caption = sImpuesto3 & " : "
       lblI3Canal4.Caption = sImpuesto3 & " : "
       lblI3Canal5.Caption = sImpuesto3 & " : "
    End If
    
    LlenaCombos
                        
    If lAlmacen Then
       txtDirecto.Visible = True
       lblDirecto.Visible = True
       txtReceta.Visible = True
       lblReceta.Visible = True
       Label(14).Visible = True
       Label(8).Visible = True
       cmdBusca(0).Visible = True
       cmdBusca(1).Visible = True
    End If
              
    'Botones
    Isql = "select tCodigoProducto, tSubGrupo, nBoton, tResumido from TPRODUCTO where nBoton > 0 order by tCodigoProducto"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    'Combos
    Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.vProducto.Descripcion, dbo.TCOMBO.nCantidad, dbo.TCOMBO.lFijo, dbo.TCOMBO.lUnico, dbo.TCOMBO.tEtiqueta, dbo.TCOMBO.nAumento,dbo.tcombo.leliminaFijo " & _
           "FROM dbo.TCOMBO LEFT OUTER JOIN dbo.vProducto ON dbo.TCOMBO.tCodigoProducto = dbo.vProducto.Codigo " & _
           "Order by tEtiqueta, dbo.vProducto.Descripcion"
    Set rsCombo = Lib.OpenRecordset(Isql, Cn)
    
    'Ofertas
    Isql = "SELECT * from tOferta"
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)
        
    'Productos x areas de impresion
    Isql = "select * from vPRODUCTOAREA order by tArea"
    Set RsGImpresion = Lib.OpenRecordset(Isql, Cn)
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
       If SAPVALOR Then
       sCodSapInf = sCodigo
        Call LLenarDatosSap
       End If
    Else
       'Cambiar el Nombre del formulario Cabecera
       ActivarBotones (True)
       
       '-- Validacion en Caso el Evento Venga del Boton de Nuevos Registros Sap
       If SAPVALOR Then
        ActivarBotones (False)
        Blanquear Me
        Iniciar
        sCodSapInf = sCodigo
        Call LLenarDatosSap
       Else
        Asignar
       End If
       
    End If
    
    If lSAP Then
        Label(21).Caption = "Codigo SAP :"
    Else
        Label(21).Caption = "Codigo de Barra :"
    End If
    If sLey1 <> "" Then
        Me.chkLey1.Visible = True
        Me.chkLey1.Caption = "Aplica Ley( " & sLey1 & " )"
    End If

'    'Configuración de la Grilla Combos
'    Call ConfGrilla(8, grdCombo, "Codigo", 2, "tCodigoProducto", 700, 2, 0, "", _
'                                 "Descripción", 2, "Descripcion", 2500, 0, 0, "", _
'                                 "Autom", 2, "lFijo", 600, 2, 4, "", _
'                                 "Cant", 2, "nCantidad", 450, 1, 0, "##0.00", _
'                                 "Elimi", 2, "lEliminaFijo", 500, 2, 4, "", _
'                                 "Unico", 2, "lUnico", 500, 2, 4, "", _
'                                 "Etiq", 2, "tEtiqueta", 800, 0, 0, "", _
'                                 "Prec.", 2, "nAumento", 600, 1, 0, "##0.00")
'    Set grdCombo.DataSource = rsCombo
'
    'Configuración de la Grilla Ofertas
    Call ConfGrilla(7, grdOferta, "Item", 2, "tOferta", 600, 2, 0, "", _
                                  "NombreOferta", 2, "tNombre", 1850, 0, 0, "", _
                                  "Hor.Ini", 2, "tHoraInicial", 700, 2, 0, "", _
                                  "Hor.Fin", 2, "tHoraFinal", 700, 2, 0, "", _
                                  "Ratio", 2, "nRatio", 550, 2, 0, "", _
                                  "Acum", 2, "lAcumulable", 600, 2, 4, "", _
                                  "Act", 2, "lActivo", 600, 2, 4, "")
    Set grdOferta.DataSource = RsOferta
    
    'Configuración de la Grilla Impresiones
    Call ConfGrilla(2, grdGrilla, "Codigo", 2, "tArea", 900, 2, 0, "", _
                                  "Descripcion", 2, "Area", 2550, 0, 0, "")
    Set grdGrilla.DataSource = RsGImpresion
    Me.tabProducto.Tab = 0
    ssTab.Tab = 0
    cmdTexto.Caption = "Registro " & frmProducto.RsCabecera.AbsolutePosition & " de " & frmProducto.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsGrupo = Nothing
    Set RsSubGrupo = Nothing
    Set RsArea = Nothing
    Set RsPrecio = Nothing
    Set RsDescargo = Nothing
    Set RsBoton = Nothing
    Set RsReceta = Nothing
    Set RsEnlace = Nothing
    Set RsMoneda = Nothing
    Set RsImpresion = Nothing
    Set RsGImpresion = Nothing
    Set CnInfhotel = Nothing
    'Cambia el Nombre del Formulario
    Set frmProductoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)

Dim xGrupo As String
Dim xRGrupoAcceso As Recordset
xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
    If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
                  '10305010 AGREGAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(0).Enabled = False
                End If
    End If
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
'    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
    cmdOpcion(4).Enabled = Activa
'    If lSAP Then
'        cmdOpcion(0).Enabled = False
'        'cmdOpcion(1).Enabled = False
'    End If
    
End Sub

'Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'   If rsCombo.RecordCount > 0 Then
'      txtQnCombo.Text = Format(rsCombo!nCantidad, "##0.00")
'      chkFijo.value = IIf(IsNull(rsCombo!lFijo), 0, Abs(Int(rsCombo!lFijo)))
'      chkUnico.value = IIf(IsNull(rsCombo!lUnico), 0, Abs(Int(rsCombo!lUnico)))
'      chkEliminaFijo.value = IIf(IsNull(rsCombo!lEliminaFijo), 0, Abs(Int(rsCombo!lEliminaFijo)))
'
'      txtEtiqueta.Text = IIf(IsNull(rsCombo!tEtiqueta), "", rsCombo!tEtiqueta)
'      txtAumPrec.Text = Format(rsCombo!nAumento, "##0.00")
'   Else
'      txtQnCombo.Text = "0.00"
'      chkFijo.value = 0
'      chkUnico.value = 0
'      chkEliminaFijo.value = 0
'      txtEtiqueta.Text = ""
'      txtAumPrec.Text = Format(0, "##0.00")
'   End If
'End Sub

Private Sub grdOferta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    AsignaDatos
End Sub

Private Sub txtCantidadItems_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtCantidadItems
End Sub


Private Sub txtCantidadItemsMinima_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtCantidadItemsMinima
End Sub
Private Sub txtDelivery_GotFocus()
   Enfoque txtDelivery
End Sub

Private Sub txtDelivery_KeyPress(KeyAscii As Integer)
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtDelivery_LostFocus()
   nPrecioDelivery = txtDelivery.Text
   CalculaDelivery
End Sub

Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtAlternativa_LostFocus()
   Call ValidaStr(txtAlternativa)
End Sub


Private Sub txtLlevar_GotFocus()
   Enfoque txtLlevar
End Sub

Private Sub txtLlevar_LostFocus()
   nPrecioLlevar = txtLlevar.Text
   CalculaLlevar
End Sub

Private Sub txtLlevar_KeyPress(KeyAscii As Integer)
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub


Private Sub txtPrecioReferencial_GotFocus()
    Enfoque txtPrecioVenta
End Sub

Private Sub txtPrecioReferencial_KeyPress(KeyAscii As Integer)
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
   If Not IsNumeric(Me.txtPrecioReferencial.Text) Then
        Me.txtPrecioReferencial.Text = "0"
   End If
End Sub

Private Sub txtPrecioVenta_GotFocus()
   Enfoque txtPrecioVenta
End Sub

Private Sub txtCanal4_GotFocus()
   Enfoque txtCanal4
End Sub

Private Sub txtCanal4_KeyPress(KeyAscii As Integer)
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtCanal4_LostFocus()
   nPrecioCanal4 = txtCanal4.Text
   CalculaCanal4
End Sub

Private Sub txtCanal5_GotFocus()
   Enfoque txtCanal5
End Sub

Private Sub txtCanal5_KeyPress(KeyAscii As Integer)
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtCanal5_LostFocus()
   nPrecioCanal5 = txtCanal5.Text
   CalculaCanal5
End Sub


Private Sub txtPrecioVenta_KeyPress(KeyAscii As Integer)
On Error GoTo fin
   PictureNumero KeyAscii
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
fin:
End Sub

Private Sub txtPrecioVenta_LostFocus()
   nPrecioVenta = txtPrecioVenta.Text
   If chkDelivery.Value = 0 Then
      nPrecioDelivery = Round(nPrecioVenta * (1 + (nDELIVERY / 100)), 2)
      txtDelivery.Text = Format(nPrecioDelivery, "###,##0." & Mid("0000000000", 1, nDecimal))
      nNetoDelivery = nPrecioVenta
      txtPrecioBaseDelivery.Caption = Format(nNetoDelivery, "###,##0." & Mid("0000000000", 1, nDecimal))
   End If
   
   If chkLlevar.Value = 0 Then
      nPrecioLlevar = Round(nPrecioVenta * (1 + (nLlevar / 100)), 2)
      txtLlevar.Text = Format(nPrecioLlevar, "###,##0." & Mid("0000000000", 1, nDecimal))
      nNetoLlevar = nPrecioVenta
      txtPrecioBaseLlevar.Caption = Format(nNetoLlevar, "###,##0.00")
   End If
   
   If chkCanal4.Value = 0 Then
      nPrecioCanal4 = Round(nPrecioVenta * (1 + (nCanal4 / 100)), 2)
      txtCanal4.Text = Format(nPrecioCanal4, "###,##0." & Mid("0000000000", 1, nDecimal))
      nNetoCanal4 = nPrecioVenta
      txtPrecioBaseCanal4.Caption = Format(nNetoCanal4, "###,##0." & Mid("0000000000", 1, nDecimal))
   End If
   
   If chkCanal5.Value = 0 Then
      nPrecioCanal5 = Round(nPrecioVenta * (1 + (nCanal5 / 100)), 2)
      txtCanal5.Text = Format(nPrecioCanal5, "###,##0." & Mid("0000000000", 1, nDecimal))
      nNetoCanal5 = nPrecioVenta
      txtPrecioBaseCanal5.Caption = Format(nNetoCanal5, "###,##0." & Mid("0000000000", 1, nDecimal))
   End If
   
   Calcula
End Sub

Private Sub Iniciar()
    txtTipoProducto.Caption = ""
    sTipoProducto = ""
    lblUsuario.Caption = ""
    lblFecha.Caption = ""
    txtTiempo.Text = "0"
'    cmdCombo(1).Enabled = False
'    cmdCombo(2).Enabled = False
    chkActivo.Value = 1
    chkDescuento.Value = 1
    chkImprime.Value = 1
 
    txtBoton.Caption = "NA"
    chkImpuesto1.Value = 0
    chkImpuesto2.Value = 0
    chkImpuesto3.Value = 0
    chkBalanza.Value = 0
    chkImpuesto1Delivery.Value = 0
    chkImpuesto2Delivery.Value = 0
    chkImpuesto3Delivery.Value = 0
    chkImpuesto1Llevar.Value = 0
    chkImpuesto2Llevar.Value = 0
    chkImpuesto3Llevar.Value = 0
    chkImpuesto1Canal4.Value = 0
    chkImpuesto2Canal4.Value = 0
    chkImpuesto3Canal4.Value = 0
    chkImpuesto1Canal5.Value = 0
    chkImpuesto2Canal5.Value = 0
    chkImpuesto3Canal5.Value = 0
    '=
    chkLCombinacion.Value = 0
    cmdCombinacion.Enabled = False
    '=
    chkDelivery.Value = 0
    chkLlevar.Value = 0
    chkCanal4.Value = 0
    chkCanal5.Value = 0
    
    nPrecio = 0
    nPrecioVenta = 0
    nPrecioDelivery = 0
    nPrecioLlevar = 0
    nPrecioCanal4 = 0
    nPrecioCanal5 = 0
    Me.chkDefineArea.Value = 0
    txtPrecioBase.Caption = "0.00"
    txtPrecioBaseDelivery.Caption = "0.00"
    txtPrecioBaseLlevar.Caption = "0.00"
    txtPrecioBaseCanal4.Caption = "0.00"
    txtPrecioBaseCanal5.Caption = "0.00"
    
    txtPrecioCosto.Text = "0.00"
    txtImpuesto1.Caption = "0.00"
    txtImpuesto2.Caption = "0.00"
    txtImpuesto3.Caption = "0.00"
    txtPrecioVenta.Text = Format(0, "###,###,##0." & Mid("0000000000", 1, nDecimal))
    txtDelivery.Text = Format(0, "###,###,##0." & Mid("0000000000", 1, nDecimal))
    txtLlevar.Text = Format(0, "###,###,##0." & Mid("0000000000", 1, nDecimal))
    txtCanal4.Text = Format(0, "###,###,##0." & Mid("0000000000", 1, nDecimal))
    txtCanal5.Text = Format(0, "###,###,##0." & Mid("0000000000", 1, nDecimal))
    
    txtInsumo.Text = "0.00"
    txtGasto.Text = "0.00"
    txtManoObra.Text = "0.00"
    txtUtilidad.Text = "0.00"
    txtPrecioCosto.Text = "0.00"
    Me.txtPrecioReferencial.Text = "0"
    txtBarra.Text = ""
    RsSubGrupo.Filter = adFilterNone
      
    chkVisualizarLocal.Value = 1
    chkVisualizarDelivery.Value = 1
    chkVisualizarLlevar.Value = 1
    chkvisualizarCanal4.Value = 1
    chkVisualizarCanal5.Value = 1
    
    'multitareas
    chkMultiArea.Value = 0
    
    chkInsumoCritico.Value = 0
    fraInsumo.Enabled = False
    scodigoInsumoCritico = ""
    lblInsumo.Caption = ""
    
    RsGImpresion.Filter = "tCodigoProducto=''"
    rsCombo.Filter = "tCodigoProducto=''"
    RsOferta.Filter = "tCodigoProducto=''"
    RsGImpresion.Requery
    rsCombo.Requery
    RsOferta.Requery
    Me.txtSKU.Text = ""
    Me.chkLey1.Value = 0
    With mscGrafico
        .chartType = VtChChartType2dPie
        .ColumnCount = 1
        .RowCount = 1
        .RowLabelCount = 0
        .ShowLegend = False
    End With
   
    For i = 1 To 20
        cmdBoton(i).backColor = vbButtonFace
        cmdBoton(i).Enabled = True
    Next i
End Sub

Private Sub Calcula()
    
    Dim nValor As Double
    nValor = 0
    
    Select Case pais ' ok
        
        Case "001" 'Bolivia
            If chkImpuesto1.Value = 1 Then
               nValor = nValor + nPorcentaje1
            End If
            If chkImpuesto2.Value = 1 Then
               nValor = nValor + nPorcentaje2
            End If
            If chkImpuesto3.Value = 1 Then
               nValor = nValor + nPorcentaje3
            End If
            nValor = (nValor / 100)
            
            nImpuesto1 = IIf(chkImpuesto1.Value = 1, (nPrecioVenta) * (nPorcentaje1 / 100), 0)
            nImpuesto2 = IIf(chkImpuesto2.Value = 1, (nPrecioVenta) * (nPorcentaje2 / 100), 0)
            nImpuesto3 = IIf(chkImpuesto3.Value = 1, (nPrecioVenta) * (nPorcentaje3 / 100), 0)
            
            nPrecio = nPrecioVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
            
            txtPrecioBase.Caption = Format(Round(nPrecio, 2), "###,###,##0.00")
            txtImpuesto1.Caption = Format(Round(nImpuesto1, 2), "###,###,##0.00")
            txtImpuesto2.Caption = Format(Round(nImpuesto2, 2), "###,###,##0.00")
            txtImpuesto3.Caption = Format(Round(nImpuesto3, 2), "###,###,##0.00")
            txtPrecioVenta.Text = Format(Round(nPrecioVenta, nDecimal), "###,###,##0." & Mid("0000000000", 1, nDecimal))
            
        Case Else 'Peru, Ecuador
            If chkImpuesto1.Value = 1 Then
               nValor = nValor + nPorcentaje1
            End If
            If chkImpuesto2.Value = 1 Then
               nValor = nValor + nPorcentaje2
            End If
            If chkImpuesto3.Value = 1 Then
               nValor = nValor + nPorcentaje3
            End If
            nValor = 1 + (nValor / 100)
            
            nImpuesto1 = IIf(chkImpuesto1.Value = 1, (nPrecioVenta / nValor) * (nPorcentaje1 / 100), 0)
            nImpuesto2 = IIf(chkImpuesto2.Value = 1, (nPrecioVenta / nValor) * (nPorcentaje2 / 100), 0)
            nImpuesto3 = IIf(chkImpuesto3.Value = 1, (nPrecioVenta / nValor) * (nPorcentaje3 / 100), 0)
            
            nPrecio = nPrecioVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
            
            txtPrecioBase.Caption = Format(Round(nPrecio, 2), "###,###,##0.00")
            txtImpuesto1.Caption = Format(Round(nImpuesto1, 2), "###,###,##0.00")
            txtImpuesto2.Caption = Format(Round(nImpuesto2, 2), "###,###,##0.00")
            txtImpuesto3.Caption = Format(Round(nImpuesto3, 2), "###,###,##0.00")
            txtPrecioVenta.Text = Format(Round(nPrecioVenta, nDecimal), "###,###,##0." & Mid("0000000000", 1, nDecimal))
        
    End Select
    
End Sub

Private Sub CalculaDelivery()
   Calcula2
End Sub

Private Sub CalculaLlevar()
   Calcula3
End Sub

Private Sub CalculaCanal4()
   Calcula4
End Sub

Private Sub CalculaCanal5()
   Calcula5
End Sub

Private Sub Calcula1()
   Dim xImpuesto As Double
   Dim nVenta As Double
   Select Case pais ' ok
   
       Case "001" 'Bolivia
            nVenta = nPrecioVenta
            xImpuesto = ((IIf(chkImpuesto1.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3.Value = 1, nPorcentaje3, 0)) / 100)
            nPrecio = nVenta - (nVenta * xImpuesto)
            txtPrecioBase.Caption = Format(Round(nPrecio, 2), "###,##0.00")
            nImpuesto1 = IIf(chkImpuesto1.Value = 1, nVenta * nPorcentaje1 / 100, 0)
            nImpuesto2 = IIf(chkImpuesto2.Value = 1, nVenta * nPorcentaje2 / 100, 0)
            nImpuesto3 = IIf(chkImpuesto3.Value = 1, nVenta * nPorcentaje3 / 100, 0)
            
            txtImpuesto1.Caption = Format(Round(nImpuesto1, 2), "###,##0.00")
            txtImpuesto2.Caption = Format(Round(nImpuesto2, 2), "###,##0.00")
            txtImpuesto3.Caption = Format(Round(nImpuesto3, 2), "###,##0.00")
            
       Case Else 'Peru, Ecuador
            nVenta = nPrecioVenta
            xImpuesto = 1 + ((IIf(chkImpuesto1.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3.Value = 1, nPorcentaje3, 0)) / 100)
            nPrecio = nVenta / xImpuesto
            txtPrecioBase.Caption = Format(Round(nPrecio, 2), "###,##0.00")
            nImpuesto1 = IIf(chkImpuesto1.Value = 1, nPrecio * nPorcentaje1 / 100, 0)
            nImpuesto2 = IIf(chkImpuesto2.Value = 1, nPrecio * nPorcentaje2 / 100, 0)
            nImpuesto3 = IIf(chkImpuesto3.Value = 1, nPrecio * nPorcentaje3 / 100, 0)
            
            txtImpuesto1.Caption = Format(Round(nImpuesto1, 2), "###,##0.00")
            txtImpuesto2.Caption = Format(Round(nImpuesto2, 2), "###,##0.00")
            txtImpuesto3.Caption = Format(Round(nImpuesto3, 2), "###,##0.00")
               
    End Select
End Sub


Private Sub Calcula2()
   Dim xImpuesto As Double
    Select Case pais ' ok
        Case "001" ' Bolivia
            xImpuesto = ((IIf(chkImpuesto1Delivery.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Delivery.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Delivery.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoDelivery = nPrecioDelivery - (nPrecioDelivery * xImpuesto)
            txtPrecioBaseDelivery.Caption = Format(Round(nNetoDelivery, 2), "###,##0.00")
            txtImpuesto1Delivery.Caption = Format(Round(IIf(chkImpuesto1Delivery.Value = 1, nPrecioDelivery * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Delivery.Caption = Format(Round(IIf(chkImpuesto2Delivery.Value = 1, nPrecioDelivery * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Delivery.Caption = Format(Round(IIf(chkImpuesto3Delivery.Value = 1, nPrecioDelivery * nPorcentaje3 / 100, 0), 2), "###,##0.00")
        
        Case Else  'Peru, Ecuador
            xImpuesto = 1 + ((IIf(chkImpuesto1Delivery.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Delivery.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Delivery.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoDelivery = nPrecioDelivery / xImpuesto
            txtPrecioBaseDelivery.Caption = Format(Round(nNetoDelivery, 2), "###,##0.00")
            txtImpuesto1Delivery.Caption = Format(Round(IIf(chkImpuesto1Delivery.Value = 1, nNetoDelivery * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Delivery.Caption = Format(Round(IIf(chkImpuesto2Delivery.Value = 1, nNetoDelivery * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Delivery.Caption = Format(Round(IIf(chkImpuesto3Delivery.Value = 1, nNetoDelivery * nPorcentaje3 / 100, 0), 2), "###,##0.00")
        
    End Select
End Sub

Private Sub Calcula3()
   Dim xImpuesto As Double
    Select Case pais 'ok
        
        Case "001" 'Bolivia
            xImpuesto = ((IIf(chkImpuesto1Llevar.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Llevar.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Llevar.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoLlevar = nPrecioLlevar - (nPrecioLlevar * xImpuesto)
            txtPrecioBaseLlevar.Caption = Format(Round(nNetoLlevar, 2), "###,##0.00")
            txtImpuesto1Llevar.Caption = Format(Round(IIf(chkImpuesto1Llevar.Value = 1, nPrecioLlevar * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Llevar.Caption = Format(Round(IIf(chkImpuesto2Llevar.Value = 1, nPrecioLlevar * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Llevar.Caption = Format(Round(IIf(chkImpuesto3Llevar.Value = 1, nPrecioLlevar * nPorcentaje3 / 100, 0), 2), "###,##0.00")
        
        Case Else 'Peru, Ecuador
            xImpuesto = 1 + ((IIf(chkImpuesto1Llevar.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Llevar.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Llevar.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoLlevar = nPrecioLlevar / xImpuesto
            txtPrecioBaseLlevar.Caption = Format(Round(nNetoLlevar, 2), "###,##0.00")
            txtImpuesto1Llevar.Caption = Format(Round(IIf(chkImpuesto1Llevar.Value = 1, nNetoLlevar * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Llevar.Caption = Format(Round(IIf(chkImpuesto2Llevar.Value = 1, nNetoLlevar * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Llevar.Caption = Format(Round(IIf(chkImpuesto3Llevar.Value = 1, nNetoLlevar * nPorcentaje3 / 100, 0), 2), "###,##0.00")
        
    End Select
End Sub

Private Sub Calcula4()
   Dim xImpuesto As Double
    Select Case pais ' ok

        Case "001" 'Bolivia
            xImpuesto = ((IIf(chkImpuesto1Canal4.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Canal4.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Canal4.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoCanal4 = nPrecioCanal4 - (nPrecioCanal4 * xImpuesto)
            txtPrecioBaseCanal4.Caption = Format(Round(nNetoCanal4, 2), "###,##0.00")
            txtImpuesto1Canal4.Caption = Format(Round(IIf(chkImpuesto1Canal4.Value = 1, nPrecioCanal4 * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Canal4.Caption = Format(Round(IIf(chkImpuesto2Canal4.Value = 1, nPrecioCanal4 * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Canal4.Caption = Format(Round(IIf(chkImpuesto3Canal4.Value = 1, nPrecioCanal4 * nPorcentaje3 / 100, 0), 2), "###,##0.00")
            
        Case Else 'Peru, Ecuador
            xImpuesto = 1 + ((IIf(chkImpuesto1Canal4.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Canal4.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Canal4.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoCanal4 = nPrecioCanal4 / xImpuesto
            txtPrecioBaseCanal4.Caption = Format(Round(nNetoCanal4, 2), "###,##0.00")
            txtImpuesto1Canal4.Caption = Format(Round(IIf(chkImpuesto1Canal4.Value = 1, nNetoCanal4 * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Canal4.Caption = Format(Round(IIf(chkImpuesto2Canal4.Value = 1, nNetoCanal4 * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Canal4.Caption = Format(Round(IIf(chkImpuesto3Canal4.Value = 1, nNetoCanal4 * nPorcentaje3 / 100, 0), 2), "###,##0.00")
    End Select
End Sub

Private Sub Calcula5()
   Dim xImpuesto As Double
    Select Case pais 'ok
        Case "001"
            xImpuesto = ((IIf(chkImpuesto1Canal5.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Canal5.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Canal5.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoCanal5 = nPrecioCanal5 - (nPrecioCanal5 * xImpuesto)
            txtPrecioBaseCanal5.Caption = Format(Round(nNetoCanal5, 2), "###,##0.00")
            txtImpuesto1Canal5.Caption = Format(Round(IIf(chkImpuesto1Canal5.Value = 1, nPrecioCanal5 * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Canal5.Caption = Format(Round(IIf(chkImpuesto2Canal5.Value = 1, nPrecioCanal5 * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Canal5.Caption = Format(Round(IIf(chkImpuesto3Canal5.Value = 1, nPrecioCanal5 * nPorcentaje3 / 100, 0), 2), "###,##0.00")
            
        Case Else
            xImpuesto = 1 + ((IIf(chkImpuesto1Canal5.Value = 1, nPorcentaje1, 0) + IIf(chkImpuesto2Canal5.Value = 1, nPorcentaje2, 0) + IIf(chkImpuesto3Canal5.Value = 1, nPorcentaje3, 0)) / 100)
            nNetoCanal5 = nPrecioCanal5 / xImpuesto
            txtPrecioBaseCanal5.Caption = Format(Round(nNetoCanal5, 2), "###,##0.00")
            txtImpuesto1Canal5.Caption = Format(Round(IIf(chkImpuesto1Canal5.Value = 1, nNetoCanal5 * nPorcentaje1 / 100, 0), 2), "###,##0.00")
            txtImpuesto2Canal5.Caption = Format(Round(IIf(chkImpuesto2Canal5.Value = 1, nNetoCanal5 * nPorcentaje2 / 100, 0), 2), "###,##0.00")
            txtImpuesto3Canal5.Caption = Format(Round(IIf(chkImpuesto3Canal5.Value = 1, nNetoCanal5 * nPorcentaje3 / 100, 0), 2), "###,##0.00")
    End Select
End Sub

Private Sub Botonera()
    RsBoton.Filter = "tSubGrupo='" & cboSubGrupo.BoundText & "'"
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 20
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
               cmdBoton(i).Caption = Trim(str(i))
            Else
               cmdBoton(i).Enabled = False
               cmdBoton(i).Caption = Trim(RsBoton!tResumido)
               If RsBoton!tCodigoProducto = txtCodigo.Text Then
                  cmdBoton(i).backColor = &H80&
                  txtBoton.Caption = Trim(str(i))
               Else
                  cmdBoton(i).backColor = &H800000
               End If
            End If
        Next i
    Else
       For i = 1 To 20
           cmdBoton(i).backColor = vbButtonFace
           cmdBoton(i).Enabled = True
           cmdBoton(i).Caption = Trim(str(i))
       Next i
    End If
End Sub

Private Sub txtRatio_LostFocus()
   txtRatio.Text = Format(txtRatio.Text, "##0.00")
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
   If txtBoton.Caption <> "NA" Then
      cmdBoton(val(txtBoton.Caption)).Caption = txtResumido.Text
   End If
End Sub

Public Sub Grafico()
    Dim xInsumo As Double
    Dim xInsumo2 As Double
    Dim xInsumo3 As Double
    Dim xInsumo4 As Double
    Dim xInsumo5 As Double
    Dim xGasto As Double
    Dim xGasto2 As Double
    Dim xGasto3 As Double
    Dim xGasto4 As Double
    Dim xGasto5 As Double
    Dim xMObra As Double
    Dim xMObra2 As Double
    Dim xMObra3 As Double
    Dim xMObra4 As Double
    Dim xMObra5 As Double
    Dim xUtilidad As Double
    Dim xUtilidad2 As Double
    Dim xUtilidad3 As Double
    Dim xUtilidad4 As Double
    Dim xUtilidad5 As Double
   
    nUtilidad = IIf(cboMoneda.BoundText = "01", nPrecio, nPrecio * nTC) - nCosto
    nUtilidad2 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioDelivery = 0, nPrecio, nNetoDelivery), IIf(nPrecioDelivery = 0, nPrecio, nPrecioDelivery) * nTC) - nCosto2
    nUtilidad3 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioLlevar = 0, nPrecio, nNetoLlevar), IIf(nPrecioLlevar = 0, nPrecio, nNetoLlevar) * nTC) - nCosto3
    nUtilidad4 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioCanal4 = 0, nPrecio, nNetoCanal4), IIf(nPrecioCanal4 = 0, nPrecio, nNetoCanal4) * nTC) - nCosto4
    nUtilidad5 = IIf(cboMoneda.BoundText = "01", IIf(nPrecioCanal5 = 0, nPrecio, nNetoCanal5), IIf(nPrecioCanal5 = 0, nPrecio, nNetoCanal5) * nTC) - nCosto5

    If nUtilidad = 0 Or nPrecio = 0 Then
       xInsumo = 0
       xGasto = 0
       xMObra = 0
       xUtilidad = 0
    Else
       xInsumo = nInsumo * 100 / (nCosto + nUtilidad)
       xGasto = nGasto * 100 / (nCosto + nUtilidad)
       xMObra = nMObra * 100 / (nCosto + nUtilidad)
       xUtilidad = nUtilidad * 100 / (nCosto + nUtilidad)
    End If
    If nUtilidad2 = 0 Or nPrecioDelivery = 0 Then
       xInsumo2 = 0
       xGasto2 = 0
       xMObra2 = 0
       xUtilidad2 = 0
    Else
       xInsumo2 = nInsumo2 * 100 / (nCosto2 + nUtilidad2)
       xGasto2 = nGasto2 * 100 / (nCosto2 + nUtilidad2)
       xMObra2 = nMObra2 * 100 / (nCosto2 + nUtilidad2)
       xUtilidad2 = nUtilidad2 * 100 / (nCosto2 + nUtilidad2)
    End If
    If nUtilidad3 = 0 Or nPrecioLlevar = 0 Then
       xInsumo3 = 0
       xGasto3 = 0
       xMObra3 = 0
       xUtilidad3 = 0
    Else
       xInsumo3 = nInsumo3 * 100 / (nCosto3 + nUtilidad3)
       xGasto3 = nGasto3 * 100 / (nCosto3 + nUtilidad3)
       xMObra3 = nMObra3 * 100 / (nCosto3 + nUtilidad3)
       xUtilidad3 = nUtilidad3 * 100 / (nCosto3 + nUtilidad3)
    End If
    If nUtilidad4 = 0 Or nPrecioCanal4 = 0 Then
       xInsumo4 = 0
       xGasto4 = 0
       xMObra4 = 0
       xUtilidad4 = 0
    Else
       xInsumo4 = nInsumo4 * 100 / (nCosto4 + nUtilidad4)
       xGasto4 = nGasto4 * 100 / (nCosto4 + nUtilidad4)
       xMObra4 = nMObra4 * 100 / (nCosto4 + nUtilidad4)
       xUtilidad4 = nUtilidad4 * 100 / (nCosto4 + nUtilidad4)
    End If
    If nUtilidad5 = 0 Or nPrecioCanal5 = 0 Then
       xInsumo5 = 0
       xGasto5 = 0
       xMObra5 = 0
       xUtilidad5 = 0
    Else
       xInsumo5 = nInsumo5 * 100 / (nCosto5 + nUtilidad5)
       xGasto5 = nGasto5 * 100 / (nCosto5 + nUtilidad5)
       xMObra5 = nMObra5 * 100 / (nCosto5 + nUtilidad5)
       xUtilidad5 = nUtilidad5 * 100 / (nCosto5 + nUtilidad5)
    End If
        
    txtInsumo.Text = Format(nInsumo, "###,###,##0.00")
    txtInsumo2.Text = Format(nInsumo2, "###,###,##0.00")
    txtInsumo3.Text = Format(nInsumo3, "###,###,##0.00")
    txtInsumo4.Text = Format(nInsumo4, "###,###,##0.00")
    txtInsumo5.Text = Format(nInsumo5, "###,###,##0.00")
    txtGasto.Text = Format(nGasto, "###,###,##0.00")
    txtGasto2.Text = Format(nGasto2, "###,###,##0.00")
    txtGasto3.Text = Format(nGasto3, "###,###,##0.00")
    txtGasto4.Text = Format(nGasto4, "###,###,##0.00")
    txtGasto5.Text = Format(nGasto5, "###,###,##0.00")
    txtManoObra.Text = Format(nMObra, "###,###,##0.00")
    txtManoObra2.Text = Format(nMObra2, "###,###,##0.00")
    txtManoObra3.Text = Format(nMObra3, "###,###,##0.00")
    txtManoObra4.Text = Format(nMObra4, "###,###,##0.00")
    txtManoObra5.Text = Format(nMObra5, "###,###,##0.00")
    txtPrecioCosto.Text = Format(nCosto, "###,###,##0.00")
    txtPrecioCosto2.Text = Format(nCosto2, "###,###,##0.00")
    txtPrecioCosto3.Text = Format(nCosto3, "###,###,##0.00")
    txtPrecioCosto4.Text = Format(nCosto4, "###,###,##0.00")
    txtPrecioCosto5.Text = Format(nCosto5, "###,###,##0.00")
    txtUtilidad.Text = Format(nUtilidad, "###,##0.00")
    txtUtilidad2.Text = Format(nUtilidad2, "###,##0.00")
    txtUtilidad3.Text = Format(nUtilidad3, "###,##0.00")
    txtUtilidad4.Text = Format(nUtilidad4, "###,##0.00")
    txtUtilidad5.Text = Format(nUtilidad5, "###,##0.00")

    With mscGrafico
         .chartType = VtChChartType2dPie
         .ColumnCount = 4
         .RowCount = 1
         .RowLabelCount = 0
         .ShowLegend = True
                      
         'Data
         'Precio de Insumo
         .Column = 1
         .ColumnLabel = "Ins.: " & Round(xInsumo) & "%"
         .Data = nInsumo
         'Gastos Generales
         .Column = 2
         .ColumnLabel = "G.G.: " & Round(xGasto) & "%"
         .Data = nGasto
         'Mano de Obra
         .Column = 3
         .ColumnLabel = "M.O.: " & Round(xMObra) & "%"
         .Data = nMObra
         'Precio Utilidad
         .Column = 4
         .ColumnLabel = "Uti.: " & Round(xUtilidad) & "%"
         .Data = nUtilidad
    End With
    
    With mscGrafico2
         .chartType = VtChChartType2dPie
         .ColumnCount = 4
         .RowCount = 1
         .RowLabelCount = 0
         .ShowLegend = True
                      
         'Data
         'Precio de Insumo
         .Column = 1
         .ColumnLabel = "Ins.: " & Round(xInsumo2) & "%"
         .Data = nInsumo
         'Gastos Generales
         .Column = 2
         .ColumnLabel = "G.G.: " & Round(xGasto2) & "%"
         .Data = nGasto
         'Mano de Obra
         .Column = 3
         .ColumnLabel = "M.O.: " & Round(xMObra2) & "%"
         .Data = nMObra
         'Precio Utilidad
         .Column = 4
         .ColumnLabel = "Uti.: " & Round(xUtilidad2) & "%"
         .Data = nUtilidad
    End With
    
    With mscGrafico3
         .chartType = VtChChartType2dPie
         .ColumnCount = 4
         .RowCount = 1
         .RowLabelCount = 0
         .ShowLegend = True
                      
         'Data
         'Precio de Insumo
         .Column = 1
         .ColumnLabel = "Ins.: " & Round(xInsumo3) & "%"
         .Data = nInsumo
         'Gastos Generales
         .Column = 2
         .ColumnLabel = "G.G.: " & Round(xGasto3) & "%"
         .Data = nGasto
         'Mano de Obra
         .Column = 3
         .ColumnLabel = "M.O.: " & Round(xMObra3) & "%"
         .Data = nMObra
         'Precio Utilidad
         .Column = 4
         .ColumnLabel = "Uti.: " & Round(xUtilidad3) & "%"
         .Data = nUtilidad
    End With
    
    With mscGrafico4
         .chartType = VtChChartType2dPie
         .ColumnCount = 4
         .RowCount = 1
         .RowLabelCount = 0
         .ShowLegend = True
                      
         'Data
         'Precio de Insumo
         .Column = 1
         .ColumnLabel = "Ins.: " & Round(xInsumo4) & "%"
         .Data = nInsumo
         'Gastos Generales
         .Column = 2
         .ColumnLabel = "G.G.: " & Round(xGasto4) & "%"
         .Data = nGasto
         'Mano de Obra
         .Column = 3
         .ColumnLabel = "M.O.: " & Round(xMObra4) & "%"
         .Data = nMObra
         'Precio Utilidad
         .Column = 4
         .ColumnLabel = "Uti.: " & Round(xUtilidad4) & "%"
         .Data = nUtilidad
    End With
    
    With mscGrafico5
         .chartType = VtChChartType2dPie
         .ColumnCount = 4
         .RowCount = 1
         .RowLabelCount = 0
         .ShowLegend = True
                      
         'Data
         'Precio de Insumo
         .Column = 1
         .ColumnLabel = "Ins.: " & Round(xInsumo5) & "%"
         .Data = nInsumo
         'Gastos Generales
         .Column = 2
         .ColumnLabel = "G.G.: " & Round(xGasto5) & "%"
         .Data = nGasto
         'Mano de Obra
         .Column = 3
         .ColumnLabel = "M.O.: " & Round(xMObra5) & "%"
         .Data = nMObra
         'Precio Utilidad
         .Column = 4
         .ColumnLabel = "Uti.: " & Round(xUtilidad5) & "%"
         .Data = nUtilidad
    End With
        
End Sub

Public Sub ProductoInfhotel()
   Dim tcuentaContable As String
   Dim rstTipoExoneracion As ADODB.Recordset
   Dim h As Integer
   
  ' tcuentaContable = Calcular("select isnull(tcuentacontable,'') as codigo from tgrupo where tcodigogrupo='" & cboGrupo.BoundText & "' ", Cn)
   tcuentaContable = Calcular("select isnull(tcuentacontable,'') as codigo from tsubgrupo where tcodigosubgrupo='" & cboSubGrupo.BoundText & "' ", Cn)
   Set rstTipoExoneracion = Lib.OpenRecordset("select tCodigoExoneracion from TEXONERACIONTIPO where ttipoproducto='" & sTipoProducto & "'", CnInfhotel)
   Dim nCorrela As String
   Dim sprefijo As String
   If Sw Then


      sprefijo = Calcular("select tValor as Codigo from vTipoProducto where Codigo='" & sTipoProducto & "' and lActivo=1", Cn)
      
      If sprefijo <> "" Then
                
                nCorrela = Calcular("select MAX(tCodigoProducto) as Codigo From TPRODUCTOPRECIO where tTipoProducto = '" & sprefijo & "'", CnInfhotel)
                
                
                If IsNull(nCorrela) Or nCorrela = "" Then
                   sCodInfhotel = sprefijo & "0001"
                Else
                   sCodInfhotel = sprefijo & Lib.Correlativo(Mid(nCorrela, 3, 4), 4)
                End If
                                           
                Isql = "insert into TPRODUCTOPRECIO( " & _
                       "tHotel, tCodigoProducto, tTipo, tTipoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                       "tDescripcionCompletaProducto, tDescripcionResumidaProducto, " & _
                       "lActivo, lTipoHabitacion, lHotel, tDescargo, lPrecioModificable, tMoneda, nPrecioCosto, " & _
                       "nPrecioVenta, tUsuario, fRegistro, fTime, tCodigoArticulo, tCodigoReceta, tOperacion, " & _
                       "nInsumo, nManoObra, nGasto,tcuentacontable) " & _
                       "values ('" & sHotel & "', '" & sCodInfhotel & "', '', '" & sprefijo & "', '', '', " & _
                              " '" & txtDetallado.Text & "', " & _
                              " '" & txtResumido.Text & "', " & _
                              "1, 0, 0, " & _
                              " '" & cboDescargo.BoundText & "', " & chkModificable.Value & ", " & _
                              " '" & cboMoneda.BoundText & "', " & nCosto & ", " & nPrecioVenta & ", " & _
                              " '" & xUsuario & "', getdate(), getdate(), '" & txtDirecto.Text & "', '" & txtReceta.Text & "', '', " & _
                              nInsumo & ", " & nMObra & ", " & nGasto & ",'" & IIf(tcuentaContable = "0", "", tcuentaContable) & "') "
                              
                CnInfhotel.Execute Isql
                If chkImpuesto1.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','01','" & sHotel & "', 1)"
                End If
                              
                If chkImpuesto2.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','02','" & sHotel & "', 1)"
                End If
                              
                If chkImpuesto3.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','03','" & sHotel & "', 1)"
                End If
                                                  
                If Not (rstTipoExoneracion.EOF Or rstTipoExoneracion.BOF) Then
                      rstTipoExoneracion.MoveFirst
                      
                      For h = 0 To rstTipoExoneracion.RecordCount - 1
                           CnInfhotel.Execute "insert into TEXONERACIONPRODUCTOPRECIO (tcodigoproducto,tcodigoexoneracion,tusuario,fregistro) values('" & sCodInfhotel & "','" & rstTipoExoneracion.Fields(0) & "','" & xUsuario & "', getdate())"
                           rstTipoExoneracion.MoveNext
                      Next h
                   
                End If
      Else
        sCodInfhotel = ""
      End If
   Else
      If sCodInfhotel <> "" Then
         Isql = "update TPRODUCTOPRECIO set " & _
                "tDescripcionCompletaProducto='" & txtDetallado.Text & "', " & _
                "tDescripcionResumidaProducto='" & txtResumido.Text & "', " & _
                "lActivo =" & chkActivo.Value & ", " & _
                "tDescargo = '" & cboDescargo.BoundText & "', " & _
                "lPrecioModificable = " & chkModificable.Value & ", " & _
                "tMoneda = '" & cboMoneda.BoundText & "', " & _
                "nPrecioVenta = " & nPrecioVenta & ", " & _
                "nPrecioCosto = " & nCosto & ", nInsumo = " & nInsumo & ", nGasto = " & nGasto & ", nManoObra = " & nMObra & ", " & _
                "tUsuario = '" & xUsuario & "', " & _
                "fRegistro = getdate(), fTime = getdate(), lReplica=1, " & _
                "tCodigoArticulo = '" & txtDirecto.Text & "', " & _
                "tCodigoReceta = '" & txtReceta.Text & "', " & _
                "tcuentacontable = '" & IIf(tcuentaContable = "0", "", tcuentaContable) & "' " & _
                " where tCodigoProducto = '" & sCodInfhotel & "'"
                
         CnInfhotel.Execute Isql
         CnInfhotel.Execute "delete from TIMPUESTOPRODUCTO where tCodigoProducto='" & sCodInfhotel & "'"
        
        If Not (rstTipoExoneracion.EOF Or rstTipoExoneracion.BOF) Then
            rstTipoExoneracion.MoveFirst
            For h = 0 To rstTipoExoneracion.RecordCount - 1
                If Calcular("select count(tcodigoproducto) as codigo  from TEXONERACIONPRODUCTOPRECIO where tcodigoproducto='" & sCodInfhotel & "'  and tcodigoexoneracion='" & rstTipoExoneracion.Fields(0) & "'", CnInfhotel) = 0 Then
                            CnInfhotel.Execute "insert into TEXONERACIONPRODUCTOPRECIO (tcodigoproducto,tcodigoexoneracion,tusuario,fregistro) values('" & sCodInfhotel & "','" & rstTipoExoneracion.Fields(0) & "','" & xUsuario & "', getdate())"
                End If
                rstTipoExoneracion.MoveNext
            Next h
        End If
        
         If chkImpuesto1.Value Then
            CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','01','" & sHotel & "', 1)"
         End If
         If chkImpuesto2.Value Then
            CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','02','" & sHotel & "', 1)"
         End If
         If chkImpuesto3.Value Then
            CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','03','" & sHotel & "', 1)"
         End If
         
      Else
      
      
      sprefijo = Calcular("select tValor as Codigo from vTipoProducto where Codigo='" & sTipoProducto & "' and lActivo=1", Cn)
      
      If sprefijo <> "" Then
                
                nCorrela = Calcular("select MAX(tCodigoProducto) as Codigo From TPRODUCTOPRECIO where tTipoProducto = '" & sprefijo & "'", CnInfhotel)
                
                
                If IsNull(nCorrela) Or nCorrela = "" Then
                   sCodInfhotel = sprefijo & "0001"
                Else
                   sCodInfhotel = sprefijo & Lib.Correlativo(Mid(nCorrela, 3, 4), 4)
                End If
                                           
                Isql = "insert into TPRODUCTOPRECIO( " & _
                       "tHotel, tCodigoProducto, tTipo, tTipoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                       "tDescripcionCompletaProducto, tDescripcionResumidaProducto, " & _
                       "lActivo, lTipoHabitacion, lHotel, tDescargo, lPrecioModificable, tMoneda, nPrecioCosto, " & _
                       "nPrecioVenta, tUsuario, fRegistro, fTime, tCodigoArticulo, tCodigoReceta, tOperacion, " & _
                       "nInsumo, nManoObra, nGasto,tcuentacontable) " & _
                       "values ('" & sHotel & "', '" & sCodInfhotel & "', '', '" & sprefijo & "', '', '', " & _
                              " '" & txtDetallado.Text & "', " & _
                              " '" & txtResumido.Text & "', " & _
                              "1, 0, 0, " & _
                              " '" & cboDescargo.BoundText & "', " & chkModificable.Value & ", " & _
                              " '" & cboMoneda.BoundText & "', " & nCosto & ", " & nPrecioVenta & ", " & _
                              " '" & xUsuario & "', getdate(), getdate(), '" & txtDirecto.Text & "', '" & txtReceta.Text & "', '', " & _
                              nInsumo & ", " & nMObra & ", " & nGasto & ",'" & IIf(tcuentaContable = "0", "", tcuentaContable) & "') "
                              
                CnInfhotel.Execute Isql
                If chkImpuesto1.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','01','" & sHotel & "', 1)"
                End If
                              
                If chkImpuesto2.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','02','" & sHotel & "', 1)"
                End If
                              
                If chkImpuesto3.Value Then
                   CnInfhotel.Execute "INSERT INTO TIMPUESTOPRODUCTO (tCodigoProducto, tCodigoImpuesto, tHotel, lAfecto) values ('" & sCodInfhotel & "','03','" & sHotel & "', 1)"
                End If
                                                  
                If Not (rstTipoExoneracion.EOF Or rstTipoExoneracion.BOF) Then
                      rstTipoExoneracion.MoveFirst
                      
                      For h = 0 To rstTipoExoneracion.RecordCount - 1
                           CnInfhotel.Execute "insert into TEXONERACIONPRODUCTOPRECIO (tcodigoproducto,tcodigoexoneracion,tusuario,fregistro) values('" & sCodInfhotel & "','" & rstTipoExoneracion.Fields(0) & "','" & xUsuario & "', getdate())"
                           rstTipoExoneracion.MoveNext
                      Next h
                   
                End If
      Else
        sCodInfhotel = ""
      End If
      
      End If
   End If
      
End Sub

Public Sub AsignaDatos()
    With RsOferta
         If .RecordCount > 0 Then
             txtNombre = !tNombre
             dtpHoraIni = Format(!tHoraInicial, "hh:mm")
             dtpHoraFin = Format(!tHorafinal, "hh:mm")
             txtRatio = Format(!nRatio, "##0.00")
             chkAcumulable = IIf(!lAcumulable = True, 1, 0)
             ChkActv = IIf(!lActivo = True, 1, 0)
             dtpFecha = IIf(IsNull(!fFecha), Date, !fFecha)
         End If
    End With
End Sub

Private Sub txtStockPlatoCritico_Change()
On Error GoTo fin:
    If Not IsNumeric(Me.txtStockPlatoCritico.Text) Then
        Me.txtStockPlatoCritico.Text = ""
    Else
        Me.txtStockPlatoCritico.Text = CInt(Me.txtStockPlatoCritico.Text)
    End If
    
    Exit Sub
fin:
    MsgBox "Error: " & error
    Me.txtStockPlatoCritico.Text = ""
End Sub

'Private Sub chkEliminaFijo_Click()
'   If Me.chkEliminaFijo.value = 1 Then
'    If Me.chkFijo.value = 0 Then
'        MsgBox "Indicar que el producto es de Cargo Automatico dentro del Combo"
'        chkEliminaFijo.value = 0
'        Me.chkFijo.SetFocus
'    End If
'   End If
'End Sub

Private Sub txtTiempo_GotFocus()
    txtTiempo.SelLength = Len(Trim(txtTiempo.Text))
End Sub

Private Sub txtTiempo_LostFocus()
    If txtTiempo.Text = "" Or Not IsNumeric(txtTiempo) Then
       txtTiempo.Text = "0"
    Else
       txtTiempo.Text = Format(txtTiempo.Text, "###,##0")
    End If
End Sub

Private Sub LLenarDatosSap()
On Error GoTo fin
    Dim sapisql As String

    
     sapisql = "exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','" & sCodigo & "','',3"

     Set rsSAPDatos = Lib.OpenRecordset(sapisql, Cn)
     With rsSAPDatos
        If .RecordCount > 0 Then
           txtDetallado = IIf(IsNull(!INF_DescripcionDetallada), "", !INF_DescripcionDetallada)
           txtResumido = IIf(IsNull(!INF_DescripcionResumida), "", !INF_DescripcionResumida)
            cboMoneda.BoundText = IIf(IsNull(!inf_moneda), "", Trim(!inf_moneda))
            chkLCombinacion = IIf(!inf_combo = True, 1, 0)
            Call chkLCombinacion_Click
            chkInsumoCritico = IIf(!inf_InsumoCritico = True, 1, 0)
            Call chkInsumoCritico_click
        Else
        
        End If
     End With
  
    Exit Sub
fin:
    MsgBox "error al procesar"
End Sub

Private Function LLenarInsumoCritico() As String
On Error GoTo fin
         'frmInsumo.RsCabecera
        Dim rsstock As Recordset
        Dim nCorrela As String
        Dim nPos As Integer
        Dim cont As Integer
        
         
           cont = Calcular("select count(*) as codigo from TINSUMO where Descripcion = UPPER('" & txtResumido.Text & "')", Cn)
          If (CInt(cont) > 0) Then
             nCorrela = Calcular("select top 1 tcodigo as Codigo from TINSUMO where Descripcion = UPPER('" & txtResumido.Text & "')", Cn)
             LLenarInsumoCritico = nCorrela
             MsgBox "Ya existe un Insumo Critico/plato con el Nombre:  " & txtResumido.Text & "", vbExclamation, sMensaje
             Exit Function
          End If
           'Obtiene el Numero de Orden
           nCorrela = Calcular("select max(tcodigo) as Codigo from TINSUMO ", Cn)
           If IsNull(nCorrela) Or nCorrela = "" Then
               nCorrela = "00000001"
               'txtCodigo.Text = nCorrela
               LLenarInsumoCritico = nCorrela
           Else
                nCorrela = Lib.Correlativo(nCorrela, 8)
                'txtCodigo.Text = nCorrela
                 LLenarInsumoCritico = nCorrela
           End If
          
         'txtResumido.Text = UCase(txtResumido.Text)
         
         
         'Llama al procedimiento para agregar un mensaje
         Dim s As String
         s = "exec usp_agregarinsumos '" & nCorrela & "','" & sUsuario & "','" & txtResumido.Text & "',0,'" & sCaja & "',1,0"
         Cn.Execute "exec usp_agregarinsumos'" & nCorrela & "','" & sUsuario & "','" & txtResumido.Text & "',0,'" & sCaja & "',1,0"
          
         Exit Function
fin:
         MsgBox "error en el porceso verificar"
End Function
