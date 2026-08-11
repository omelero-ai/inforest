VERSION 5.00
Begin VB.Form frmImpresoraDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10410
   Icon            =   "frmImpresoraDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10410
   Begin VB.Frame FrameConfImpresion 
      Caption         =   "Configuración Tamaño de Letra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtF12 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   2760
         TabIndex        =   71
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtF11 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   2760
         TabIndex        =   70
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2400
         Picture         =   "frmImpresoraDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1920
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
         Height          =   495
         Index           =   5
         Left            =   2400
         Picture         =   "frmImpresoraDetalle.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   2520
         Width           =   1170
      End
      Begin VB.TextBox txtF10 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   2760
         TabIndex        =   64
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtF9 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   960
         TabIndex        =   62
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtF8 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   960
         TabIndex        =   52
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtF7 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   51
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtF6 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   50
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtF5 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   49
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF4 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   48
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtF3 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   47
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtF2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   46
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtF1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   45
         Top             =   360
         Width           =   855
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
         Height          =   495
         Index           =   4
         Left            =   2400
         Picture         =   "frmImpresoraDetalle.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label Label16 
         Caption         =   "Linea 12 :"
         Height          =   255
         Left            =   2040
         TabIndex        =   72
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Linea 11 :"
         Height          =   255
         Left            =   2040
         TabIndex        =   69
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Linea 10 :"
         Height          =   255
         Left            =   2040
         TabIndex        =   63
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Linea 9 :"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Linea 8 :"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Linea 7 :"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Linea 6 :"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Linea 5 :"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Linea 4 :"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Linea 3 :"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Linea 2 :"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Linea 1 :"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame frmVineta 
      Caption         =   "Configuraciones Impresora Viñeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5220
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   10365
      Begin VB.CommandButton cmdRegresar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   150
         Picture         =   "frmImpresoraDetalle.frx":0F98
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Regresar"
         Top             =   3900
         Width           =   990
      End
      Begin VB.TextBox txtAltoVineta 
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
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   36
         Text            =   " "
         Top             =   450
         Width           =   2220
      End
      Begin VB.TextBox txtAnchoVineta 
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
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   37
         Text            =   " "
         Top             =   975
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(mm= milímetros)"
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
         Left            =   4650
         TabIndex        =   42
         Top             =   975
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(mm= milímetros)"
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
         Left            =   4650
         TabIndex        =   41
         Top             =   525
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Alto Vineta : "
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
         Left            =   1005
         TabIndex        =   40
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Vineta : "
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
         Left            =   810
         TabIndex        =   38
         Top             =   975
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10350
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4530
      Width           =   10410
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
         Left            =   9120
         Picture         =   "frmImpresoraDetalle.frx":13DA
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   7950
         Picture         =   "frmImpresoraDetalle.frx":14CC
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   5610
         Picture         =   "frmImpresoraDetalle.frx":15CE
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   6780
         Picture         =   "frmImpresoraDetalle.frx":1B00
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5430
         TabIndex        =   12
         Top             =   60
         Width           =   5490
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmImpresoraDetalle.frx":2032
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmImpresoraDetalle.frx":2574
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmImpresoraDetalle.frx":2AB6
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4950
            Picture         =   "frmImpresoraDetalle.frx":2FF8
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4470
            Picture         =   "frmImpresoraDetalle.frx":353A
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3990
            Picture         =   "frmImpresoraDetalle.frx":3A7C
            Style           =   1  'Graphical
            TabIndex        =   13
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
            Left            =   1470
            TabIndex        =   24
            Top             =   150
            Width           =   2445
         End
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
      Height          =   4440
      Left            =   1890
      TabIndex        =   19
      Top             =   0
      Width           =   8460
      Begin VB.Frame frmImpresoraFiscal 
         Caption         =   "Impresora Fiscal"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   6840
         TabIndex        =   73
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         Begin VB.ComboBox cboBitPorSegundo 
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
            ItemData        =   "frmImpresoraDetalle.frx":3FBE
            Left            =   120
            List            =   "frmImpresoraDetalle.frx":3FF8
            TabIndex        =   77
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtPuertoSerial 
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
            Left            =   120
            MaxLength       =   5
            TabIndex        =   75
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "BIt por Segundo:"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Puerto Serial:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdFuenteImpresion 
         Caption         =   "Configuracion de Fuentes"
         Height          =   375
         Left            =   3840
         TabIndex        =   65
         Top             =   3960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdConfigurar 
         Caption         =   "Configuraciones de Viñetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3810
         TabIndex        =   34
         ToolTipText     =   "Configuraciones de Impresora de Viñeta"
         Top             =   3450
         Width           =   2865
      End
      Begin VB.CheckBox chkVine 
         Alignment       =   1  'Right Justify
         Caption         =   "Impresora de Viñetas :"
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
         Height          =   390
         Left            =   930
         TabIndex        =   33
         Top             =   3450
         Width           =   2340
      End
      Begin VB.TextBox txtFSPedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5985
         MaxLength       =   5
         TabIndex        =   32
         Text            =   " "
         Top             =   3080
         Width           =   700
      End
      Begin VB.TextBox txtFSPrecuenta 
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
         Left            =   3810
         MaxLength       =   5
         TabIndex        =   31
         Text            =   " "
         Top             =   3080
         Width           =   700
      End
      Begin VB.TextBox txtNS 
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
         Left            =   2910
         MaxLength       =   20
         TabIndex        =   6
         Text            =   " "
         Top             =   2580
         Width           =   3780
      End
      Begin VB.ComboBox cboFuente2 
         Height          =   315
         ItemData        =   "frmImpresoraDetalle.frx":4069
         Left            =   2910
         List            =   "frmImpresoraDetalle.frx":406B
         TabIndex        =   5
         Text            =   "cboFuente2"
         Top             =   2190
         Width           =   3780
      End
      Begin VB.ComboBox cboFuente1 
         Height          =   315
         ItemData        =   "frmImpresoraDetalle.frx":406D
         Left            =   2910
         List            =   "frmImpresoraDetalle.frx":406F
         TabIndex        =   4
         Text            =   "cboFuente1"
         Top             =   1815
         Width           =   3780
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   2910
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   711
         Width           =   5235
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
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1035
      End
      Begin VB.ComboBox cboRuta 
         Height          =   315
         ItemData        =   "frmImpresoraDetalle.frx":4071
         Left            =   2910
         List            =   "frmImpresoraDetalle.frx":4073
         TabIndex        =   2
         Text            =   "cboRuta"
         Top             =   1062
         Width           =   3780
      End
      Begin VB.ComboBox cboFuente 
         Height          =   315
         ItemData        =   "frmImpresoraDetalle.frx":4075
         Left            =   2910
         List            =   "frmImpresoraDetalle.frx":4077
         TabIndex        =   3
         Text            =   "cboFuente"
         Top             =   1440
         Width           =   3780
      End
      Begin VB.Label lblfuenteImpresion 
         Alignment       =   1  'Right Justify
         Caption         =   "Tamaño de Fuente (Impresion de Envio de Pedido  y Precuenta )"
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
         Height          =   375
         Left            =   480
         TabIndex        =   66
         Top             =   3960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Impresión Pedido :"
         Height          =   195
         Index           =   9
         Left            =   4650
         TabIndex        =   30
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Precuentas :"
         Height          =   195
         Index           =   8
         Left            =   2850
         TabIndex        =   29
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tamaño de Fuente para Impresiones en Mobile Inforest"
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
         Height          =   420
         Index           =   3
         Left            =   -30
         TabIndex        =   28
         Top             =   2955
         Width           =   2820
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Número de Serie de la Impresora :"
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   27
         Top             =   2640
         Width           =   2400
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fuente (Ancho) de Impresión Pedido :"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   26
         Top             =   2250
         Width           =   2670
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fuente (Normal) Impresión Pedido :"
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   25
         Top             =   1875
         Width           =   2475
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   23
         Top             =   750
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   22
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Impresora :"
         Height          =   195
         Index           =   2
         Left            =   2025
         TabIndex        =   21
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Impresión de Documentos :"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   20
         Top             =   1485
         Width           =   2475
      End
   End
   Begin VB.Image Image 
      Height          =   4455
      Left            =   45
      Picture         =   "frmImpresoraDetalle.frx":4079
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmImpresoraDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjPrinter As Object
Dim p As Object
'Dim Puerto(100) As String

Sub LlenaCombos()
    Dim i As Integer
    i = 0
    For Each ObjPrinter In Printers
        If TypeOf ObjPrinter Is Printer Then
            'cboRuta.AddItem ObjPrinter.DeviceName & " en " & ObjPrinter.Port
            cboRuta.AddItem ObjPrinter.DeviceName
            'Puerto(i) = ObjPrinter.Port
            i = i + 1
        End If
    Next
End Sub

Sub Asignar()
    If frmImpresora.RsCabecera.RecordCount = 0 Then
       txtCodigo = ""
       txtDescripcion.Text = ""
       'txtRuta.Text = ""
       cboRuta.Text = ""
       cboFuente.Text = ""
       cboFuente1.Text = ""
       cboFuente2.Text = ""
       Me.chkVine.value = 0
       Me.txtAltoVineta.Text = ""
       Me.txtAnchoVineta.Text = ""
       txtFSPedido.Text = ""
       txtFSPrecuenta.Text = ""
       Me.txtPuertoSerial.Text = ""
       Me.cboBitPorSegundo.Text = ""
    Else
    
       With frmImpresora.RsCabecera
          Dim i As Integer

          'Cuadro de Texto
          txtCodigo = IIf(IsNull(!timpresora), "", !timpresora)
          txtDescripcion = IIf(IsNull(!tDescripcion), "", !tDescripcion)
          txtNS.Text = IIf(IsNull(!tNumeroSerie), "", !tNumeroSerie)
          
          txtFSPrecuenta.Text = IIf(IsNull(!nFontSizePrecuenta), "0", !nFontSizePrecuenta)

          txtFSPedido.Text = IIf(IsNull(!nFontSizeEnvio), "0", !nFontSizeEnvio)
            
            
          'vineta
          Me.txtAltoVineta.Text = IIf(IsNull(!nHeightVineta), "0", !nHeightVineta)
          Me.txtAnchoVineta.Text = IIf(IsNull(!nWidthVineta), "0", !nWidthVineta)
           Me.chkVine = IIf(!lvineta = True, 1, 0)
           If chkVine.value = 1 Then
            Me.cmdConfigurar.Enabled = True
           Else
           Me.cmdConfigurar.Enabled = False
           End If
            
          'Data Combo
          'txtRuta.Text = IIf(IsNull(!tRuta), "", !tRuta)
          cboRuta.Text = IIf(IsNull(!tDevice), "", !tDevice)
          
          'Fuentes
          LlenaFuente (frmImpresora.RsCabecera!tDevice)
          cboFuente.Text = IIf(IsNull(!tFont), "", !tFont)
          cboFuente1.Text = IIf(IsNull(!tFont1), "", !tFont1)
          cboFuente2.Text = IIf(IsNull(!tFont2), "", !tFont2)
          
          Me.txtPuertoSerial.Text = IIf(IsNull(!tPuertoSerial), "", !tPuertoSerial)
          Me.cboBitPorSegundo.Text = IIf(IsNull(!tBitPorSegundo), "", !tBitPorSegundo)
          
       End With
    End If
    
    If pais = "003" Then
        frmImpresoraFiscal.Visible = True
    End If
    
    
    cmdTexto.Caption = "Registro " & IIf(frmImpresora.RsCabecera.RecordCount = 0, 0, frmImpresora.RsCabecera.AbsolutePosition) & " de " & frmImpresora.RsCabecera.RecordCount
        
End Sub

Private Sub cboRuta_Click()
'  txtRuta.Text = Puerto(cboRuta.ListIndex)
  LlenaFuente (cboRuta.Text)
End Sub

Private Sub chkVine_Click()
If chkVine.value = 1 Then
    Me.cmdConfigurar.Enabled = True
Else
    Me.cmdConfigurar.Enabled = False
End If
End Sub

Private Sub cmdConfigurar_Click()

frmVineta.Left = 0
frmVineta.Height = 4740
frmVineta.Top = 0

frmVineta.Visible = True
Me.txtAltoVineta.SetFocus
End Sub

Private Sub cmdFuenteImpresion_Click()
FrameConfImpresion.Visible = True
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmImpresora.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmImpresora.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmImpresora.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmImpresora.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmImpresora.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmImpresora.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmImpresora.RsCabecera.RecordCount = 0, 0, frmImpresora.RsCabecera.AbsolutePosition) & " de " & frmImpresora.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               txtDescripcion.SetFocus
               cboFuente.Clear
                         
          Case Is = 1 'Grabar
               Dim nCorrela As String
          
               'ChequeaDatos
               If txtDescripcion.Text = "" Then MsgBox "Ingrese la Descripción de la Impresora", vbExclamation, sMensaje: txtDescripcion.SetFocus: Exit Sub
               If cboRuta.Text = "" Then MsgBox "Ingrese el Nombre de la Impresora", vbExclamation, sMensaje: cboRuta.SetFocus: Exit Sub
                     
                     
               If Me.chkVine.value = 1 Then
                    If Me.txtAltoVineta.Text = "" Or Me.txtAltoVineta.Text = "0" Then
                        MsgBox "Ingrese un tamaño válido de Viñeta", vbCritical, sMensaje
                        Exit Sub
                    End If
                       If Me.txtAnchoVineta.Text = "" Or Me.txtAnchoVineta.Text = "0" Then
                        MsgBox "Ingrese un tamaño válido de Viñeta", vbCritical, sMensaje
                        Exit Sub
                    End If
               End If
                     
                     
                     
                     
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tImpresora) as Codigo from TIMPRESORA where tCaja='" & sCaja & "'", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
   
                  'Cambiar el SQL
                  Isql = "insert into TIMPRESORA( tImpresora, tCaja, tDescripcion, tDevice, tFont, tFont1, tFont2, tUsuario, tNumeroSerie, fRegistro, nFontSizePrecuenta, nFontSizeEnvio, lvineta,nWidthVineta, nHeightVineta, tPuertoSerial, tBitPorSegundo ) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & sCaja & "', " & _
                                " '" & txtDescripcion.Text & "', " & _
                                " '" & cboRuta.Text & "', " & _
                                " '" & cboFuente.Text & "', " & _
                                " '" & cboFuente1.Text & "', " & _
                                " '" & cboFuente2.Text & "', " & _
                                " '" & sUsuario & "', " & _
                                " '" & txtNS.Text & "', " & _
                                " getdate() ," & Val(txtFSPrecuenta.Text) & ",'" & Val(txtFSPedido.Text) & "', " & Me.chkVine.value & ", " & Val(Me.txtAnchoVineta.Text) & ", " & Val(Me.txtAltoVineta.Text) & ", '" & Trim(Me.txtPuertoSerial.Text) & "','" & Trim(Me.cboBitPorSegundo.Text) & "')"
                                
                  Cn.Execute Isql
                  frmImpresora.RsCabecera.Sort = "tImpresora ASC"
                  frmImpresora.RsCabecera.Requery
                  frmImpresora.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmImpresora.RsCabecera.RecordCount = 0, 0, frmImpresora.RsCabecera.AbsolutePosition) & " de " & frmImpresora.RsCabecera.RecordCount
                  
               Else
                  'Cambiar el SQL
                  Isql = "update TIMPRESORA set " & _
                         "tDescripcion ='" & txtDescripcion.Text & "', " & _
                         "tDevice ='" & cboRuta.Text & "', " & _
                         "tFont ='" & cboFuente.Text & "',nFontSizePrecuenta=" & Val(txtFSPrecuenta.Text) & ",nfontsizeenvio=" & Val(txtFSPedido.Text) & ",  " & _
                         "tFont1 ='" & cboFuente1.Text & "', " & _
                         "tFont2 ='" & cboFuente2.Text & "', " & _
                         "tNumeroSerie ='" & txtNS.Text & "', " & _
                         " lvineta=" & Me.chkVine.value & ",nWidthVineta=" & Val(Me.txtAnchoVineta.Text) & ", nHeightVineta=" & Val(Me.txtAltoVineta.Text) & " , tpuertoserial='" & Trim(Me.txtPuertoSerial.Text) & "', tBitPorSegundo='" & Trim(Me.cboBitPorSegundo.Text) & "' " & _
                         "where tImpresora = '" & txtCodigo & "' and tCaja ='" & sCaja & "'"
                       
                   Cn.Execute Isql
                   nPos = frmImpresora.RsCabecera.Bookmark
                   frmImpresora.RsCabecera.Requery
                   If frmImpresora.RsCabecera.RecordCount = 0 Then
                      frmImpresora.RsCabecera.Filter = adFilterNone
                   End If
                   frmImpresora.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 'Eliminar
          
               If frmImpresora.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar la Impresora " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               Cn.Execute "delete from TIMPRESORA where tImpresora = '" & txtCodigo & "' and tCaja = '" & sCaja & "'"
               frmImpresora.RsCabecera.Requery
               
               If frmImpresora.RsCabecera.RecordCount <> 0 Then
                  frmImpresora.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmImpresora.RsCabecera.RecordCount = 0, 0, frmImpresora.RsCabecera.AbsolutePosition) & " de " & frmImpresora.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          Case Is = 3 'Salir
               Unload Me
               
          Case Is = 4 ' oculta el frame de ingreso de tamaño de letras para impresionde Precuenta y envio de pedido.
               FrameConfImpresion.Visible = False
          Case Is = 5 ' grabar los tamaños de letras de las diferentes lineas de Precuenta y envio de pedido.
                Dim F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12 As Double

                If txtF1.Text = "" Or Not IsNumeric(txtF1.Text) Then
                    F1 = 0
                Else
                   F1 = txtF1.Text
                End If

                If txtF2.Text = "" Or Not IsNumeric(txtF2.Text) Then
                    F2 = 0
                Else
                    F2 = txtF2.Text
                End If
                If txtF3.Text = "" Or Not IsNumeric(txtF3.Text) Then
                    F3 = 0
                Else
                    F3 = txtF3.Text
                End If
                If txtF4.Text = "" Or Not IsNumeric(txtF4.Text) Then
                    F4 = 0
                Else
                    F4 = txtF4.Text
                End If
                If txtF5.Text = "" Or Not IsNumeric(txtF5.Text) Then
                    F5 = 0
                Else
                    F5 = txtF5.Text
                End If
                If txtF6.Text = "" Or Not IsNumeric(txtF6.Text) Then
                    F6 = 0
                Else
                    F6 = txtF6.Text
                End If
                If txtF7.Text = "" Or Not IsNumeric(txtF7.Text) Then
                    F7 = 0
                Else
                    F7 = txtF7.Text
                End If
                 If txtF8.Text = "" Or Not IsNumeric(txtF8.Text) Then
                    F8 = 0
                Else
                    F8 = txtF8.Text
                End If
                 If txtF9.Text = "" Or Not IsNumeric(txtF9.Text) Then
                    F9 = 0
                Else
                    F9 = txtF9.Text
                End If
                 If txtF10.Text = "" Or Not IsNumeric(txtF10.Text) Then
                    F10 = 0
                Else
                    F10 = txtF10.Text
                End If
                 If txtF11.Text = "" Or Not IsNumeric(txtF11.Text) Then
                    F11 = 0
                Else
                    F11 = txtF11.Text
                End If
                 If txtF12.Text = "" Or Not IsNumeric(txtF12.Text) Then
                    F12 = 0
                Else
                    F12 = txtF12.Text
                End If
            
                  
                 Dim Impresora As Recordset
                 
                Isql = "select * from TIMPRESORAIMPRESION where timpresora='" & txtCodigo.Text & "'"
                 Set Impresora = Lib.OpenRecordset(Isql, Cn)
                
                If Impresora.RecordCount = 0 Then
                    Isql = "INSERT INTO TIMPRESORAIMPRESION VALUES('" & txtCodigo.Text & "'," & F1 & "," & F2 & "," & F3 & "," & F4 & "," & F5 & "," & F6 & "," & F7 & "," & F8 & "," & F9 & "," & F10 & "," & F11 & "," & F12 & ")"
                    Cn.Execute Isql
                    MsgBox ("Se proceso Correctamente")
                Else
                    Isql = "update  TIMPRESORAIMPRESION set tf1 = " & F1 & ", tf2 = " & F2 & ", tf3 = " & F3 & ", tf4 = " & F4 & ", tf5 = " & F5 & ", tf6 = " & F6 & ", tf7 = " & F7 & ", tf8 = " & F8 & ", tf9 = " & F9 & ", tf10 = " & F10 & ", tf11 = " & F11 & ", tf12 = " & F12 & " where timpresora = '" & txtCodigo.Text & "'"
                    Cn.Execute Isql
                    MsgBox ("Se proceso Correctamente")
                End If
   End Select

End Sub
Private Sub llenarFuente()
On Error GoTo fin
    Dim IFuente As Recordset
    If txtCodigo.Text <> "" Then
        Isql = "select * from TIMPRESORAIMPRESION where timpresora='" & txtCodigo.Text & "'"
        Set IFuente = Lib.OpenRecordset(Isql, Cn)
        If IFuente.RecordCount > 0 Then
            txtF1.Text = IFuente!tF1
            txtF2.Text = IFuente!tF2
            txtF3.Text = IFuente!tF3
            txtF4.Text = IFuente!tF4
            txtF5.Text = IFuente!tF5
            txtF6.Text = IFuente!tF6
            txtF7.Text = IFuente!tF7
            txtF8.Text = IFuente!tF8
            txtF9.Text = IFuente!tF9
            txtF10.Text = IFuente!tF10
            txtF11.Text = IFuente!tF11
            txtF12.Text = IFuente!tF12
        
        End If
    End If
Exit Sub
fin:
MsgBox ("error al cargar")
End Sub
 
Private Sub cmdPreview_Click()

   'If Not Imprimir("004") Then
       'Exit Sub
   'End If
   
    Dim X As Printer
    For Each X In Printers
        If X.DeviceName = "FinePrint" Then
            Set Printer = X
           ' MsgBox "IMPRESORA ENCONTRADA"
        End If
    Next
   
   Dim RsImpresion As Recordset
   Dim sTf1, sTf2, sTf3, sTf4, sTf5, sTf6, sTf7, sTf8, sTf9, sTf10, sTf11, sTf12 As Double
   Set RsImpresion = Lib.OpenRecordset("select *  from TIMPRESORAIMPRESION where timpresora='" & txtCodigo.Text & "'", Cn)
    
    If RsImpresion.RecordCount > 0 Then
    sTf1 = RsImpresion!tF1
    sTf2 = RsImpresion!tF2
    sTf3 = RsImpresion!tF3
    sTf4 = RsImpresion!tF4
    sTf5 = RsImpresion!tF5
    sTf6 = RsImpresion!tF6
    sTf7 = RsImpresion!tF7
    sTf8 = RsImpresion!tF8
    sTf9 = RsImpresion!tF9
    sTf10 = RsImpresion!tF10
    sTf11 = RsImpresion!tF11
    sTf12 = RsImpresion!tF12
    
    End If
    
    
    
 
'    '----------Cabecera-----------------------------------------------------------------
  Printer.Font.Size = sTf1
   Printer.Print ""
   

   Printer.FontName = sFont2
   Printer.FontBold = False

    Printer.Print "Orden :  0001 - LINEA 1"  ' & RsImpresion!orden

    
    Printer.Font.Size = sTf2
    
    Printer.Print "Hora  : " & Format(Time, "HH:mm:ss") & " - LINEA 2 "
   Printer.Print " "

    Printer.Font.Size = sTf3

    Printer.Print "Mesa  : Mesa-001 - LINEA 3" ' & IIf(IsNull(RsImpresion!Mesa), "", Trim(RsImpresion!Mesa)), 17, 0
    Printer.Print " "
    Printer.FontName = sFont1
    Printer.FontBold = False
    Printer.Font.Size = sTf4
    Printer.Print "Pedido    : 0000000000001 - LINEA 4" '& sPedido
    Printer.Font.Size = sTf5
    Printer.Print "Fecha    :  " & FechaServidor() & " - LINEA 5"
    Printer.FontBold = False
    Printer.Font.Size = sTf6
    Printer.Print "Area     : Area Prueba - LINEA 6" '& sArea
    Printer.Font.Size = sTf7
    Printer.FontBold = False
    Printer.Print "Mesero ó Motoriz.  : MOZO01 Prueba - LINEA 7" '& RsImpresion!Mozo
    Printer.Font.Size = sTf8
    Printer.Print "T/Pedido : En el Local Prueba -  LINEA 8" '& sBoton1 & " " & IIf(RsImpresion!Prioridad, "Urgente", "Regular"), 23, 11
'
    Printer.Font.Size = sTf9
    Printer.Print "Personas : 100 - LINEA 9 " '& RsImpresion!nAdulto

    Printer.Font.Size = sTf10
    Printer.Print "Observac.: XXXXXXXXXXXXXXXXX - LINEA 10" '& IIf(IsNull(RsImpresion!Observacion), "", Trim(RsImpresion!Observacion)), 23, 11
'

    Printer.Print "--------------- XXXXXXXXX --------------"
    Printer.Font.Size = sTf11
    Printer.Print " DETALLES DEL PRODUCTO -  LINEA 11 "
'
             Printer.Print String(48, "-")
    Printer.Font.Size = sTf12
   Printer.Print "Caja     : CAJA 0001 Prueba - LINEA 12 " '& sCaja & " - " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
'
        Printer.Print " "


   Printer.Print
   Printer.EndDoc

Exit Sub
fin:
MsgBox ("error al procesar el preview")
End Sub

Private Sub cmdRegresar_Click()
Me.frmVineta.Visible = False
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Impresoras "
    fraDetalle.Caption = Me.Caption
    
    'Llena todos los Combos
    LlenaCombos
   
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmImpresora.RsCabecera.Find ("tImpresora = '" & frmImpresora.RsCabecera!tImpresora & "'")
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmImpresora.RsCabecera.RecordCount = 0, 0, frmImpresora.RsCabecera.AbsolutePosition) & " de " & frmImpresora.RsCabecera.RecordCount
     Call llenarFuente
    If pais = "003" Then
        frmImpresoraFiscal.Visible = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImpresoraDetalle = Nothing
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

Public Sub LlenaFuente(sImpresora As String)
    Dim i As Integer
    For Each p In Printers
        'If UCase(P.Port) = UCase(sImpresora) Then
        If UCase(p.DeviceName) = UCase(sImpresora) Then
           Set Printer = p
           cboFuente.Clear
           Printer.Orientation = vbPRORPortrait
           For i = 0 To Printer.FontCount - 1       ' Determina el número de fuentes.
               cboFuente.AddItem Printer.Fonts(i)
               cboFuente1.AddItem Printer.Fonts(i)
               cboFuente2.AddItem Printer.Fonts(i)
           Next i
           Exit For
        End If
    Next p
End Sub

Private Sub txtAltoVineta_KeyPress(KeyAscii As Integer)
 TabNext KeyAscii
   Numerico KeyAscii, txtAltoVineta
End Sub

Private Sub txtAnchoVineta_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF1_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF2_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF3_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF4_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF5_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF6_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF7_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF8_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF9_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF10_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF11_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtF12_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtAnchoVineta
End Sub
Private Sub txtFSPedido_KeyPress(KeyAscii As Integer)
  TabNext KeyAscii
   Numerico KeyAscii, txtFSPedido
End Sub

Private Sub txtFSPrecuenta_KeyPress(KeyAscii As Integer)
 TabNext KeyAscii
   Numerico KeyAscii, txtFSPrecuenta
End Sub
