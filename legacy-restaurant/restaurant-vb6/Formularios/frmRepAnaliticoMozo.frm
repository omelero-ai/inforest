VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepAnaliticoMozo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analítico de Productos por Mesero"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmRepAnaliticoMozo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
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
      Left            =   1290
      Picture         =   "frmRepAnaliticoMozo.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7545
      Width           =   1260
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
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
      Left            =   2550
      Picture         =   "frmRepAnaliticoMozo.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7545
      Width           =   1350
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
      Index           =   2
      Left            =   5235
      Picture         =   "frmRepAnaliticoMozo.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7545
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
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
      Left            =   0
      TabIndex        =   36
      Top             =   -30
      Width           =   7815
      Begin VB.CheckBox ChkFranjaHoraria 
         Caption         =   "Franja Horaria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   58
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Frame Frame6 
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
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   7560
         Begin VB.CheckBox chkUsuario 
            Caption         =   "Todos los Usuarios"
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
            Left            =   5030
            TabIndex        =   57
            Top             =   600
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.CheckBox chkMozo 
            Caption         =   "Todos los Meseros"
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
            Left            =   5030
            TabIndex        =   55
            Top             =   240
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.OptionButton optOP 
            Caption         =   "Usuario"
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
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   660
            Width           =   1215
         End
         Begin VB.OptionButton optOP 
            Caption         =   "Mesero"
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
            TabIndex        =   52
            Top             =   270
            Value           =   -1  'True
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo cboMozo 
            Height          =   315
            Left            =   1740
            TabIndex        =   54
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSDataListLib.DataCombo cboUsuario 
            Height          =   315
            Left            =   1740
            TabIndex        =   56
            Top             =   600
            Width           =   3135
            _ExtentX        =   5530
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
      End
      Begin VB.Frame Frame5 
         Caption         =   " Valores "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         TabIndex        =   50
         Top             =   5145
         Width           =   1680
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
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
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
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
            Index           =   1
            Left            =   90
            TabIndex        =   20
            Top             =   585
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
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
            Index           =   2
            Left            =   90
            TabIndex        =   21
            Top             =   900
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Origen de Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   5400
         TabIndex        =   49
         Top             =   5145
         Width           =   2325
         Begin VB.OptionButton optOrigen 
            Caption         =   "Producción (a+b+c)"
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
            Index           =   0
            Left            =   135
            TabIndex        =   26
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(a) Ventas"
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
            Left            =   135
            TabIndex        =   27
            Top             =   572
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(b) Cortesías"
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
            Index           =   2
            Left            =   135
            TabIndex        =   28
            Top             =   844
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(c) Ctas. Corrientes"
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
            Index           =   3
            Left            =   135
            TabIndex        =   29
            Top             =   1116
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   31
            Top             =   1660
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   32
            Top             =   1935
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Pedidos Facturados"
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
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   30
            Top             =   1388
            Width           =   2025
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Comensales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   1920
         TabIndex        =   48
         Top             =   5145
         Width           =   1440
         Begin VB.OptionButton optComenzales 
            Caption         =   "Todos"
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
            Index           =   2
            Left            =   225
            TabIndex        =   24
            Top             =   915
            Width           =   930
         End
         Begin VB.OptionButton optComenzales 
            Caption         =   "Adultos"
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
            Left            =   225
            TabIndex        =   22
            Top             =   285
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optComenzales 
            Caption         =   "Niños"
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
            Index           =   1
            Left            =   225
            TabIndex        =   23
            Top             =   600
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   3480
         TabIndex        =   47
         Top             =   5145
         Width           =   1800
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado"
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
            Index           =   1
            Left            =   210
            TabIndex        =   33
            Top             =   675
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
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
            Left            =   210
            TabIndex        =   25
            Top             =   330
            Value           =   -1  'True
            Width           =   1545
         End
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
         Height          =   375
         Index           =   0
         Left            =   4245
         Picture         =   "frmRepAnaliticoMozo.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3045
         Width           =   765
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
         Height          =   375
         Index           =   1
         Left            =   4245
         Picture         =   "frmRepAnaliticoMozo.frx":0C64
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3435
         Width           =   765
      End
      Begin VB.TextBox txtProducto 
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
         Left            =   1875
         TabIndex        =   8
         Top             =   3075
         Width           =   2265
      End
      Begin VB.TextBox txtTurno 
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
         Left            =   1875
         TabIndex        =   11
         Top             =   3465
         Width           =   2265
      End
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos los SubGrupos"
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
         Left            =   5130
         TabIndex        =   7
         Top             =   2685
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Todos los Tipos"
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
         Left            =   5130
         TabIndex        =   1
         Top             =   390
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
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
         Left            =   5130
         TabIndex        =   13
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todos los Grupos"
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
         Left            =   5130
         TabIndex        =   5
         Top             =   2295
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Canales de Venta"
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
         Left            =   5130
         TabIndex        =   3
         Top             =   785
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Productos"
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
         Left            =   5130
         TabIndex        =   10
         Top             =   3090
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1875
         TabIndex        =   16
         Top             =   4305
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1875
         TabIndex        =   14
         Top             =   3870
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   3855
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   4290
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1875
         TabIndex        =   0
         Top             =   375
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   1875
         TabIndex        =   2
         Top             =   770
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   1875
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1890
         TabIndex        =   6
         Top             =   2670
         Width           =   3135
         _ExtentX        =   5530
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
      Begin VB.CheckBox chkCFacturados 
         Caption         =   "Valorizar con Cero la Pre Venta Facturada"
         Height          =   405
         Left            =   1800
         TabIndex        =   18
         Top             =   4680
         Width           =   4635
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubGrupo :"
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
         Height          =   195
         Index           =   6
         Left            =   765
         TabIndex        =   46
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Producto :"
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
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   427
         Width           =   1500
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         Height          =   210
         Index           =   2
         Left            =   1155
         TabIndex        =   44
         Top             =   3510
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
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
         Height          =   210
         Left            =   660
         TabIndex        =   43
         Top             =   3945
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
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
         Height          =   210
         Left            =   750
         TabIndex        =   42
         Top             =   4365
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
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
         Height          =   210
         Index           =   4
         Left            =   1140
         TabIndex        =   41
         Top             =   2325
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Canla de Venta :"
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
         Height          =   210
         Index           =   5
         Left            =   420
         TabIndex        =   40
         Top             =   825
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
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
         Height          =   210
         Index           =   7
         Left            =   900
         TabIndex        =   39
         Top             =   3120
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar"
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
      Left            =   3885
      Picture         =   "frmRepAnaliticoMozo.frx":0D66
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7545
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   960
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepAnaliticoMozo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim RsTipo As Recordset
Dim RsTipoPedido As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsMozo As Recordset
Dim Rsturno As Recordset
Dim Detallado As New dsrAnaliticoMozo
Dim Consolidado As New dsrAnaliticoMozoC
Dim sCriterio As String
Dim xCriterio As String
Dim sTitulo As String
Dim sCriterioM As String
Dim sProducto As String
Dim sTurno As String
Dim fInicio As Date
Dim fFinal As Date

Dim tMozoCajero As String

Sub LlenaCombos()
    With cboTipoProducto
         Isql = "Select * from vTipoProducto"
         Set RsTipo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
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
        
    With cboMozo
         Isql = "Select * from vMozo where substring(Codigo,1,1)<>'*' and lactivo=1 order by descripcion"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUsuario
         Isql = "Select * from TUSUARIO where substring(tCodigoUsuario,1,1)<>'*' and lactivo=1 order by tdetallado"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "tDetallado"
             .ListField = "tDetallado"
             .BoundColumn = "tCodigoUsuario"
    End With
    
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
            
End Sub

Private Sub cboGrupo_Change()
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
    End With
    
    sProducto = ""
    txtProducto.Text = ""
End Sub



Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   If Index = 0 Then
      xCriterio = "lActivo = 1 "
      If cboGrupo.Text <> "" Then
         xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
      End If
      
      If cboSubGrupo.Text <> "" Then
         xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
      End If
      
      Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where " & xCriterio & " order by Descripcion"
      frmBusca.nPredeterm = 3
      Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                             "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                             "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                             "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                             "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
      frmBusca.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      sProducto = sCodigo
      txtProducto.Text = sDescrip
   Else
      Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
      frmBusca.nPredeterm = 2
      Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
                                             "Caja", 2, "tCaja", 1100, 0, 0, "", _
                                             "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
                                             "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
      frmBusca.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      sTurno = sCodigo
      txtTurno.Text = sCodigo
   End If

End Sub

Private Sub chkMozo_Click()
If chkMozo.Value = 1 Then
   cboMozo.Enabled = False
   cboMozo.Text = ""
   tMozoCajero = "M"
Else
   cboMozo.Enabled = True
   tMozoCajero = "M"
End If
End Sub

Private Sub chkproducto_Click()
   If chkProducto.Value = 1 Then
      sProducto = ""
      txtProducto.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.Value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub

Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.Text = ""
      cboGrupo.Enabled = False
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkSubGrupo_Click()
   If chkSubGrupo.Value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub

Private Sub chkTipo_Click()
   If chkTipo.Value = 1 Then
      cboTipoProducto.Enabled = False
      cboTipoProducto.Text = ""
   Else
      cboTipoProducto.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
        Unload Me
        Exit Sub
    End If
   
    sCriterio = ""
    xCriterio = ""
    sCriterioM = ""
    sTitulo = ""
   
    If chkTurno.Value = 0 Then
        If sTurno = "" Then
           MsgBox "Debe escoger un Turno", vbCritical, sMensaje
           Exit Sub
        End If
      
        Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
        If Rsturno.RecordCount > 0 Then
            sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
            
        Else
            MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
        End If
   Else
        If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
            MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
            Exit Sub
            dtpFecFin.SetFocus
        End If
        
            If ChkFranjaHoraria.Value = 0 Then
                sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
            Else
                sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " Franja Horaria:  " & Format(dtpHorIni.Value, "HH:mm") & " - " & Format(dtpHorFin.Value, "HH:mm")
            End If
                
        
        'sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
    
   End If
      
  
  
    If chkTipo.Value = 0 Then
        If cboTipoProducto.Text = "" Then
            MsgBox "Debe escoger un Tipo de Producto", vbCritical, sMensaje
            Exit Sub
        End If
     
    End If
   
    If chkMozo.Value = 0 Then
        If cboMozo.Text = "" Then
            MsgBox "Debe escoger un Mesero", vbCritical, sMensaje
            Exit Sub
        End If
        
    End If
    
    If chkUsuario.Value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "Debe escoger un Usuario", vbCritical, sMensaje
         Exit Sub
      End If
    End If
     
    If chkTipoPedido.Value = 0 Then
        If cboTipoPedido.Text = "" Then
            MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
            Exit Sub
        End If
    
    End If
      
    If chkGrupo.Value = 0 Then
        If cboGrupo.Text = "" Then
            MsgBox "Debe escoger el grupo", vbCritical, sMensaje
            Exit Sub
        End If
     
    End If
   
    If chkSubGrupo.Value = 0 Then
        If cboSubGrupo.Text = "" Then
            MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
            Exit Sub
        End If
    
    End If
   
    If chkProducto.Value = 0 Then
        If sProducto = "" Then
            MsgBox "Debe escoger el producto", vbCritical, sMensaje
            Exit Sub
        End If
    
    End If
   
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If rsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = True
            If optOpcion(0).Value Then
                Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detallado.PaperOrientation = crPortrait
            Else
                Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Consolidado.PaperOrientation = crPortrait
            End If
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
               
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            If optOpcion(0).Value Then
                Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detallado.PaperOrientation = crPortrait
                Detallado.PrintOut
            Else
                Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Consolidado.PaperOrientation = crPortrait
                Consolidado.PrintOut
            End If
               
        Case Is = 2 ' Salir
            Unload Me
          
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            If optOpcion(0).Value Then
                Detallado.ExportOptions.FormatType = 21
                Detallado.ExportOptions.DestinationType = 1
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                If cmdSave.FileName = "" Then
                   Exit Sub
                End If
                Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                Detallado.Export False
            Else
                Consolidado.ExportOptions.FormatType = 21
                Consolidado.ExportOptions.DestinationType = 1
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                If cmdSave.FileName = "" Then
                   Exit Sub
                End If
                Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
                Consolidado.Export False
            End If
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub ChkTurno_Click()
   If chkTurno.Value = 1 Then
      cmdBusca(1).Enabled = False
      txtTurno.Text = ""
      sTurno = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
       ChkFranjaHoraria.Enabled = True
      ChkFranjaHoraria.Value = 0
   Else
      cmdBusca(1).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
       ChkFranjaHoraria.Enabled = False
      ChkFranjaHoraria.Value = 0
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    
    Isql = "Select * from MTURNO order by tTurno DESC"
    Set Rsturno = Lib.OpenRecordset(Isql, Cn)
    
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    cboTipoProducto.Enabled = False
    cboTipoProducto.Text = ""
    cboTipoPedido.Enabled = False
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    cboMozo.Enabled = False
    cboMozo.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsSubgrupo = Nothing
   Set RsTipo = Nothing
   Set RsMozo = Nothing
End Sub

Public Sub Genera()
    Dim oComando As clsComando
    Dim sComensales As String
    Dim sPrecio As String
    Dim tipopedido As String
    Screen.MousePointer = vbHourglass
    sComensales = ""
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_AnaliticoMozo", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    If chkCFacturados.Value = 0 Then
                If Me.optOrigen(4).Value = True Then ' combinacion
                    If optValor(0).Value = True Then ' venta
                         sPrecio = "dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad"
                    ElseIf optValor(1).Value = True Then ' neto
                        sPrecio = "dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
                    ElseIf optValor(2).Value = True Then ' costo
                    sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
                    End If
                Else
                    If optValor(0).Value = True Then '  venta
                            sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                    ElseIf optValor(1).Value = True Then ' neto
                            sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                    ElseIf optValor(2).Value = True Then ' costo
                    sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                    End If
                End If
    Else
                If Me.optOrigen(4).Value = True Then ' combinacion
                    If optValor(0).Value = True Then ' venta
                         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad end "
                    ElseIf optValor(1).Value = True Then ' neto
                        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad end "
                    ElseIf optValor(2).Value = True Then ' costo
                    sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad end "
                    End If
                Else
                    If optValor(0).Value = True Then '  venta
                            sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                    ElseIf optValor(1).Value = True Then ' neto
                            sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end"
                    ElseIf optValor(2).Value = True Then ' costo
                    sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end"
                    End If
                End If
    End If
    
    If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.Value = 0 Then
    
        tipopedido = cboTipoPedido.BoundText
    Else
        tipopedido = ""
    End If
    
    
    
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")

    
    
    sComensales = IIf(optComenzales(0).Value, " isnull(A.Adultos,0) as nComenzales ", IIf(optComenzales(1).Value, " isnull(A.Ninos,0) as nComenzales ", " isnull(A.Adultos,0)+isnull(A.Ninos,0) as nComenzales "))
    oComando.CreateParameter "@flagFranjaHoraria", adBoolean, adParamInput, 1, IIf(ChkFranjaHoraria.Value = 1, True, False)
    oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).Value
    oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).Value
    oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).Value
    oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).Value
    oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).Value
    oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).Value
    oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).Value
    oComando.CreateParameter "@flagOpcion", adBoolean, adParamInput, 1, optOpcion(0).Value
    oComando.CreateParameter "@flagTurnoOFecha", adBoolean, adParamInput, 1, chkTurno.Value
    oComando.CreateParameter "@tTipoProducto", adVarChar, adParamInput, 30, cboTipoProducto.Text
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 30, cboMozo.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@turno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@comensales", adVarChar, adParamInput, 200, sComensales
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 150, sPrecio
    oComando.CreateParameter "@tMozoCajero", adVarChar, adParamInput, 2, tMozoCajero
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 50, cboUsuario.Text
         
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    If optOpcion(0).Value Then
        Detallado.DiscardSavedData
        Detallado.Database.SetDataSource rsReporte
        
        Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
        Detallado.Text10.SetText "Analítico de Productos por Meseros de " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))
        Detallado.Text18.SetText localConectado
        Detallado.Text13.SetText sRazonSocial
        Detallado.Text9.SetText IIf(optValor(0).Value = True, "Venta", IIf(optValor(1).Value = True, "Neto", "Costo"))
        
        frmEmite.CRViewer.ReportSource = Detallado
    Else
        Consolidado.DiscardSavedData
        Consolidado.Database.SetDataSource rsReporte
        Consolidado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
        Consolidado.Text10.SetText "Analítico de Productos por Meseros de " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))
        Consolidado.Text5.SetText localConectado
        Consolidado.Text13.SetText sRazonSocial
        Consolidado.Text4.SetText IIf(optValor(0).Value = True, "Total Venta:", IIf(optValor(1).Value = True, "Total Neto:", "Total Costo:"))
        frmEmite.CRViewer.ReportSource = Consolidado
    End If
End Sub

Private Sub optOP_Click(Index As Integer)
    If optOP(0).Value Then  'MOZO
        chkMozo.Enabled = True
        
        cboUsuario.Enabled = False
        chkUsuario.Enabled = False
        chkUsuario.Value = 1
        tMozoCajero = "M"
    Else                    'USUARIO
        chkUsuario.Enabled = True
        
        cboMozo.Enabled = False
        chkMozo.Enabled = False
        chkMozo.Value = 1
        tMozoCajero = "U"
    End If
End Sub

Private Sub chkUsuario_Click()
    If chkUsuario.Value = 1 Then
       cboUsuario.Enabled = False
       cboUsuario.Text = ""
       tMozoCajero = "U"
    Else
       cboUsuario.Enabled = True
       tMozoCajero = "U"
    End If
End Sub















