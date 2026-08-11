VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepVentaTipoProductoIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Por Tipo de Producto"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12795
   Icon            =   "frmRepVentaTipoProductoIntegrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   30
      TabIndex        =   37
      Top             =   0
      Width           =   12735
      Begin VB.Frame Frame7 
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
         Height          =   1200
         Left            =   120
         TabIndex        =   49
         Top             =   4320
         Width           =   1785
         Begin VB.OptionButton optTipo 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1545
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   1425
         End
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Todos los Tipos"
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
         Left            =   5610
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todos los Grupos"
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
         Left            =   5610
         TabIndex        =   10
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Pedidos"
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
         Left            =   5610
         TabIndex        =   8
         Top             =   645
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   5610
         TabIndex        =   12
         Top             =   1455
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Agrupación"
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
         Left            =   3720
         TabIndex        =   42
         Top             =   2880
         Width           =   1815
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   25
            Top             =   660
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   24
            Top             =   330
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1365
         Left            =   2040
         TabIndex        =   41
         Top             =   2880
         Width           =   1545
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   990
            Width           =   1335
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   330
            Value           =   -1  'True
            Width           =   1380
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
         Height          =   2280
         Left            =   5640
         TabIndex        =   40
         Top             =   2880
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
      Begin VB.Frame Frame5 
         Caption         =   "Servidores En Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5715
         Left            =   8160
         TabIndex        =   39
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar Todo"
            Height          =   360
            Left            =   2880
            TabIndex        =   36
            Top             =   5280
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   5055
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   8916
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.CheckBox chkFHoraria 
         Caption         =   "Franja Horaria"
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
         Left            =   5640
         TabIndex        =   17
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Frame Frame6 
         Caption         =   "Visualización"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   1815
         Begin VB.OptionButton optRanking 
            Caption         =   "Ambos"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   990
            Width           =   1185
         End
         Begin VB.OptionButton optRanking 
            Caption         =   "Por Cantidad"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton optRanking 
            Caption         =   "Por Montos"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   2100
         TabIndex        =   15
         Top             =   2265
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
         Format          =   50987009
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   2100
         TabIndex        =   13
         Top             =   1845
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
         Format          =   50987009
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4050
         TabIndex        =   14
         Top             =   1860
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   50987011
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4050
         TabIndex        =   16
         Top             =   2265
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   50987011
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   2100
         TabIndex        =   5
         Top             =   225
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   2100
         TabIndex        =   7
         Top             =   615
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   2100
         TabIndex        =   9
         Top             =   1020
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   1410
         Width           =   3375
         _ExtentX        =   5953
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Producto :"
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
         Index           =   1
         Left            =   420
         TabIndex        =   48
         Top             =   285
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
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
         Left            =   795
         TabIndex        =   47
         Top             =   1950
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
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
         Left            =   900
         TabIndex        =   46
         Top             =   2355
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
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
         Index           =   4
         Left            =   1380
         TabIndex        =   45
         Top             =   1095
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pedido :"
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
         Index           =   5
         Left            =   600
         TabIndex        =   44
         Top             =   690
         Width           =   1425
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
         Left            =   1050
         TabIndex        =   43
         Top             =   1500
         Width           =   975
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
      Left            =   5670
      Picture         =   "frmRepVentaTipoProductoIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5820
      Width           =   1455
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
      Left            =   8580
      Picture         =   "frmRepVentaTipoProductoIntegrado.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5820
      Width           =   1455
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
      Left            =   4215
      Picture         =   "frmRepVentaTipoProductoIntegrado.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5820
      Width           =   1455
   End
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
      Left            =   2760
      Picture         =   "frmRepVentaTipoProductoIntegrado.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5820
      Width           =   1455
   End
   Begin VB.CommandButton CmdActualiza 
      Caption         =   "Actualizar"
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
      Left            =   7125
      Picture         =   "frmRepVentaTipoProductoIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5820
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   2760
      Top             =   5910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepVentaTipoProductoIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim RsTempoE As Recordset
Dim RsTempoT As Recordset
Dim RsTempoC As Recordset
Dim RsTempoP As Recordset
Dim RsTurno As Recordset
Dim RsSalon As Recordset
Dim RsTipoPedido As Recordset
Dim sector(500) As String
Dim RsTipo As Recordset
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim Resumidol As New dsrTipoProductoVentaIntegrado
Dim sCriterio As String
Dim xCriterio As String
Dim sTitulo As String
Dim sTexto As String
Dim fInicio As Date
Dim fFinal As Date

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
           
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
End Sub

Private Sub cboGrupo_Click(Area As Integer)
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
    End With
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub

Private Sub ChkGrupo_Click()
   If chkGrupo.value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkSubGrupo_Click()
   If chkSubGrupo.value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub

Private Sub chkTipo_Click()
   If chkTipo.value = 1 Then
      cboTipoProducto.Enabled = False
      cboTipoProducto.Text = ""
   Else
      cboTipoProducto.Enabled = True
   End If
End Sub

Private Sub chkTodo_Click()
 activaCheck IIf(chkTodo.value = 1, True, False)
End Sub
Public Sub activaCheck(Tipo As Boolean)
    Dim o As Integer
    For o = 1 To lstServidores.ListItems.Count
        If Tipo = True Then
            If lstServidores.ListItems.Item(o).SubItems(3) <> "Desconectado" Then

                lstServidores.ListItems.Item(o).Checked = Tipo
            End If
        Else
            lstServidores.ListItems.Item(o).Checked = Tipo
        End If
    Next o
       ' validaCantidadSeleccion
End Sub
Private Sub CmdActualiza_Click()
If MsgBox("Esta operación puede tardar varios minutos" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
    Screen.MousePointer = vbHourglass
    configuraListaServidores
    Me.chkTodo.value = 0
    MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
    Dim p, r As Integer
     Cn.Execute " delete from " & sNombre
   sCriterio = ""
   xCriterio = ""
   sTitulo = ""
  r = 0
   p = 0
   For p = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(p).Checked = True Then
            r = r + 1
        End If
   Next p
   If r = 0 Then
      MsgBox "Debe seleccionar al menos un Item de la Lista de Servidores", vbCritical, sMensaje
   Exit Sub
   End If
  
      If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "
      xCriterio = " MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "

   
   If chkTipo.value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de Producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and TipoProducto ='" & cboTipoProducto.Text & "'"
   End If
   
   If chkTipoPedido.value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.BoundText & "'"
   End If
      
   If chkGrupo.value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and Grupo='" & cboGrupo.Text & "'"
   End If
    
   If chkSubGrupo.value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and subGrupo='" & cboSubGrupo.Text & "'"
   End If
     
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
             
     
                        frmEmite.CRViewer.DisplayGroupTree = True
                        Resumidol.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Resumidol.PaperOrientation = crLandscape
     
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
                        Resumidol.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Resumidol.PaperOrientation = crLandscape
                        Resumidol.PrintOut
                         
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
       
                        Resumidol.ExportOptions.FormatType = 21
                        Resumidol.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                                Exit Sub
                        End If
                        Resumidol.ExportOptions.DiskFileName = cmdSave.FileName
                        Resumidol.Export False
       
               
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub


Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    sNombre = ""
    
    sNombre = dbTemporal(sCaja, 9, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "TipoProducto", "nvarchar(50)", "Producto", "nvarchar(200)", "cantidad", "float", "venta", "float", "neto", "float", "costo", "float")
    dtpFecIni.value = Date
    dtpFecFin.value = Date
   
    sTurno = ""
    cboTipoProducto.Enabled = False
    cboTipoProducto.Text = ""
    
    cboTipoPedido.Enabled = False
   
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    'configuraListaServidores
    obtieneListaServidores
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsTempoE = Nothing
   Set RsTempoT = Nothing
   Set RsTempoC = Nothing
   Set RsTempoP = Nothing
   Set RsTurno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsGrupo = Nothing
   Set RsTipo = Nothing
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
On Error GoTo sigue
   Dim sCombo As String
   Dim sOrden As String
   Dim sPrecio As String
   Dim sNeto As String
   Dim sCosto As String
   Dim conectado As Boolean
   Dim sCostoPropiedad As String
   Dim sCostoComboPropiedad As String
   Dim w, z As Long
   Dim cs As String
   Dim cl As String
   Dim sector As String
   Dim llocal As String
   Dim valo As Double
  
   Dim RsDatos As New ADODB.Recordset
   
   Dim oComando As clsComando
   Screen.MousePointer = vbHourglass
  
    valo = 0
    
   'Produccion
   If optOrigen(0).value = True Then
      If optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      ElseIf optValor(2).value Then
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) *  dbo.DPEDIDO.nCantidad"
      Else
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
      End If
         sNeto = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) *  dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
   End If
   
   'Venta
   If optOrigen(1).value = True Then
      If optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      ElseIf optValor(2).value = True Then
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
      Else
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         
      End If
        sNeto = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Costo"
      
   End If
   
   'Cortesia
   If optOrigen(2).value = True Then
      If optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      ElseIf optValor(2).value Then
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
      Else
         sPrecio = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
      End If
         sNeto = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
         sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      
   End If

   'Cuentas Corrientes
   If optOrigen(3).value = True Then
     If optValor(0).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Venta"
     ElseIf optValor(1).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Neto"
     ElseIf optValor(2).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
     Else
        sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
     End If
        sNeto = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Costo"
     
   End If
   
   'Combinacion
   If optOrigen(4).value = True Then
      If optValor(0).value Then
         sPrecio = "dbo.CPEDIDO.nVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.CPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      ElseIf optValor(2).value Then
         sPrecio = "(isnull(dbo.vProducto.nInsumo,0)+isnull(dbo.vProducto.nGasto,0)+isnull(dbo.vProducto.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
      Else
         sPrecio = "dbo.CPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
      End If
        sNeto = "dbo.CPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
        sCosto = "(isnull(dbo.vProducto.nInsumo,0)+isnull(dbo.vProducto.nGasto,0)+isnull(dbo.vProducto.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
        sTexto = "Precio Costo"
        
   End If
            
   'Cargos
   If optOrigen(5).value = True Then
     If optValor(0).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Venta"
     ElseIf optValor(1).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Neto"
     ElseIf optValor(2).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
     Else
        sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
     End If
        sNeto = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Costo"
     
   End If
   
   'Pedidos Facturados
   If optOrigen(6).value = True Then
      If optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      ElseIf optValor(2).value Then
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
      Else
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
      End If
         sNeto = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If

   sOrden = ""

   
        If Me.chkFHoraria.value = 1 Then
            Dim dias, t As Integer
            dias = Me.dtpFecFin.value - Me.dtpFecIni.value
            For t = 0 To dias
                   fInicio = Format(dtpFecIni.value + t, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
                   fFinal = Format(dtpFecIni.value + t, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
                    For w = 1 To lstServidores.ListItems.Count
                        If lstServidores.ListItems.Item(w).Checked = True Then
                                 conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                                    If conectado = False Then
                                        MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                                        lstServidores.ListItems.Item(w).Checked = False
                                        GoTo sigue1
                                    End If
                                    Set oComando = New clsComando
                                    If Not oComando.CreateCmdSp("spRep_TipoProductoVentaIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                                            Set oComando = Nothing
                                            Exit Sub
                                    End If
                                    oComando.CreateParameter "@flagPropiedades", adBoolean, adParamInput, 1, False
                                    oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).value
                                    oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).value
                                    oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).value
                                    oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).value
                                    oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).value
                                    oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).value
                                    oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).value
                                    oComando.CreateParameter "@flagOpcion", adBoolean, adParamInput, 1, optOpcion(0).value
                                    oComando.CreateParameter "@flagVVenta", adBoolean, adParamInput, 1, optValor(0).value
                                    oComando.CreateParameter "@flagVNeto", adBoolean, adParamInput, 1, optValor(1).value
                                    oComando.CreateParameter "@flagVCosto", adBoolean, adParamInput, 1, optValor(2).value
                                    oComando.CreateParameter "@tTipoProducto", adVarChar, adParamInput, 30, cboTipoProducto.Text
                                    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, cboTipoPedido.BoundText
                                    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
                                    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
                                    oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 300, sPrecio
                                    oComando.CreateParameter "@sNeto", adVarChar, adParamInput, 300, sNeto
                                    oComando.CreateParameter "@sCosto", adVarChar, adParamInput, 300, sCosto
                                    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 200, sOrden
                                    oComando.CreateParameter "@sBoton2", adVarChar, adParamInput, 30, sBoton2
                                    oComando.CreateParameter "@sBoton3", adVarChar, adParamInput, 30, sBoton3
                                    oComando.CreateParameter "@sBoton4", adVarChar, adParamInput, 30, sBoton4
                                    oComando.CreateParameter "@sBoton5", adVarChar, adParamInput, 30, sBoton5
                                    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                                    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                                    If Not oComando.GetParamOK Then
                                       Set oComando = Nothing
                                       Exit Sub
                                    End If
                                    
                                   ' cs = LTrim(Mid(lstServidores.ListItems.Item(w), 1, 2))
                                    cl = lstServidores.ListItems.Item(w)
                                    sector = lstServidores.ListItems.Item(w).SubItems(1)
                                    llocal = lstServidores.ListItems.Item(w).SubItems(2)
                                    Set RsDatos = oComando.GetSP()
                                    If Not (RsDatos.EOF Or RsDatos.BOF) Then
                                        RsDatos.MoveFirst
                                            For z = 0 To RsDatos.RecordCount - 1
                                                 '   valo = IIf(IsNull(RsDatos.Fields(1)), 0, RsDatos.Fields(1))
                                                    Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "'," & RsDatos.Fields(2) & "," & RsDatos.Fields(3) & " ," & RsDatos.Fields(4) & "," & RsDatos.Fields(5) & ")"
                                                      'Debug.Print Isql
                                                    Cn.Execute Isql
                                                RsDatos.MoveNext
                                            Next z
                                     End If
                    End If
sigue1:
            Next w
            Next t
        Else
           fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
           fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
        For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spRep_TipoProductoVentaIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                        Set oComando = Nothing
                        Exit Sub
                End If
                        oComando.CreateParameter "@flagPropiedades", adBoolean, adParamInput, 1, False
                        oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).value
                        oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).value
                        oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).value
                        oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).value
                        oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).value
                        oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).value
                        oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).value
                        oComando.CreateParameter "@flagOpcion", adBoolean, adParamInput, 1, optOpcion(0).value
                        oComando.CreateParameter "@flagVVenta", adBoolean, adParamInput, 1, optValor(0).value
                        oComando.CreateParameter "@flagVNeto", adBoolean, adParamInput, 1, optValor(1).value
                        oComando.CreateParameter "@flagVCosto", adBoolean, adParamInput, 1, optValor(2).value
                        oComando.CreateParameter "@tTipoProducto", adVarChar, adParamInput, 30, cboTipoProducto.Text
                        oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, cboTipoPedido.BoundText
                        oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
                        oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
                        oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 300, sPrecio
                        oComando.CreateParameter "@sNeto", adVarChar, adParamInput, 300, sNeto
                        oComando.CreateParameter "@sCosto", adVarChar, adParamInput, 300, sCosto
                        oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 200, sOrden
                        oComando.CreateParameter "@sBoton2", adVarChar, adParamInput, 30, sBoton2
                        oComando.CreateParameter "@sBoton3", adVarChar, adParamInput, 30, sBoton3
                        oComando.CreateParameter "@sBoton4", adVarChar, adParamInput, 30, sBoton4
                        oComando.CreateParameter "@sBoton5", adVarChar, adParamInput, 30, sBoton5
                        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                        If Not oComando.GetParamOK Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                        
                      '  cs = LTrim(Mid(lstServidores.ListItems.Item(w), 1, 2))
                        cl = lstServidores.ListItems.Item(w)
                        sector = lstServidores.ListItems.Item(w).SubItems(1)
                        llocal = lstServidores.ListItems.Item(w).SubItems(2)
                        Set RsDatos = oComando.GetSP()
                        If Not (RsDatos.EOF Or RsDatos.BOF) Then
                            RsDatos.MoveFirst
                                For z = 0 To RsDatos.RecordCount - 1
                                   
                                       ' valo = IIf(IsNull(RsDatos.Fields(1)), 0, RsDatos.Fields(1))
                                        'Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "'," & Val(valo) & ",'" & Format(RsDatos.Fields(0), "mm/DD/yyyy") & "' )"
                                            Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "'," & RsDatos.Fields(2) & "," & RsDatos.Fields(3) & "," & RsDatos.Fields(4) & "," & RsDatos.Fields(5) & "  )"
                                            Cn.Execute Isql
                                   
                                    RsDatos.MoveNext
                                Next z
                         End If
       End If
sigue:
       Next w
       End If
       
        ' seleccionado en el local
        If optOpcion(1).value = True Then ' local
               
               
                ' ranking por monto
                If Me.optRanking(1).value = True Then
                    
                    ' valorizado por costo
                    If optValor(2).value = True Then
                    
                        If optTipo(0).value = True Then ' detallado
                            Isql = "select local , producto, sum(costo) as venta  from " & sNombre & " group by local,producto "
                            
                        Else ' resumido
                        
                            Isql = "select local , tipoproducto, sum(costo) as venta  from " & sNombre & " group by local,tipoproducto"
                            
                        End If
                    ' valorizado venta / neto
                    Else
                        If optTipo(0).value = True Then
                            Isql = "select local , producto, sum(venta) as venta  from " & sNombre & " group by local,producto "
                        Else
                            Isql = "select local , tipoproducto, sum(venta) as venta  from " & sNombre & " group by local,tipoproducto "
                        End If
                    End If
                    
                     ' ranking por monto
                ElseIf Me.optRanking(0).value = True Then ' cantidad
                
                        If optTipo(0).value = True Then
                        
                            Isql = "select local , producto, sum(cantidad) as venta  from " & sNombre & " group by local,producto "
                        Else
                            Isql = "select local , tipoproducto, sum(cantidad) as venta  from " & sNombre & " group by local,tipoproducto "
                        End If
                    
                Else ' ambos
                    
                    If optTipo(0).value = True Then ' detallao
                        If optValor(2).value = True Then ' costo
                            Isql = "select local+' Cant' , producto, sum(cantidad) as venta  from " & sNombre & " group by local,producto  union select local+' Mont', producto, sum(costo) as venta  from " & sNombre & " group by local,producto "
                        Else ' neto /venta
                            Isql = "select local+' Cant' , producto, sum(cantidad) as venta  from " & sNombre & " group by local,producto  union select local+' Mont', producto, sum(venta) as venta  from " & sNombre & " group by local,producto "
                        End If
                        
                    Else ' resumido
                        If optValor(2).value = True Then ' costo
                            Isql = "select local+' Cant' , tipoproducto, sum(cantidad) as venta  from " & sNombre & " group by local,tipoproducto  union select local+' Mont', tipoproducto, sum(costo) as venta  from " & sNombre & " group by local,tipoproducto "
                        Else ' neto /venta
                            Isql = "select local+' Cant' , tipoproducto, sum(cantidad) as venta  from " & sNombre & " group by local,tipoproducto  union select local+' Mont', tipoproducto, sum(venta) as venta  from " & sNombre & " group by local,tipoproducto "
                        End If
                        
                    End If
                End If
 
               Set rsReporte = Lib.OpenRecordset(Isql, Cn)
                Resumidol.DiscardSavedData
                Resumidol.Database.SetDataSource rsReporte
                Resumidol.PaperOrientation = crLandscape
                
                Resumidol.ReportTitle = "Por Locales. Consolidado entre: " & Me.dtpFecIni.value & "  y " & Me.dtpFecFin.value
                If Me.chkFHoraria.value = 1 Then
                    Resumidol.ReportTitle = Resumidol.ReportTitle & " Franja horaria: " & Format(Me.dtpHorIni.value, "HH:MM") & " - " & Format(Me.dtpHorFin.value, "HH:MM")
                End If
                Resumidol.Text20.SetText sRazonSocial
                Resumidol.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos") '& " (Monto Expresado en Soles)"
                Resumidol.Text6.SetText "Paloteo Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Valorización: " & IIf(optRanking(0).value = True, "Cantidad", IIf(optValor(0).value = True, optValor(0).Caption & " (Montos Expresados en Soles)", IIf(optValor(1).value = True, optValor(1).Caption & " (Montos Expresados en Soles)", optValor(2).Caption & " (Montos Expresados en Soles)")))
                frmEmite.CRViewer.ReportSource = Resumidol
           
           End If
   
            If optOpcion(0).value = True Then '  sector
                
                ' ranking por monto
                If Me.optRanking(1).value = True Then
                    
                    ' valorizado por costo
                    If optValor(2).value = True Then
                    
                        If optTipo(0).value = True Then ' detallado
                            Isql = "select sector as local, producto, sum(costo) as venta from " & sNombre & " group by sector,producto "
                        Else ' resumido
                            Isql = "select sector as local, tipoproducto, sum(costo) as venta from " & sNombre & " group by sector,tipoproducto "
                        End If
                    ' valorizado venta / neto
                    Else
                        If optTipo(0).value = True Then
                            Isql = "select sector as local, producto, sum(venta) as venta from " & sNombre & " group by sector,producto "
                        Else
                            Isql = "select sector as local, tipoproducto, sum(venta) as venta from " & sNombre & " group by sector,tipoproducto "
                        End If
                    End If
                    
                     ' ranking por monto
                ElseIf Me.optRanking(0).value = True Then ' cantidad
                        If optTipo(0).value = True Then
                            Isql = "select sector as local, producto, sum(cantidad) as venta from " & sNombre & " group by sector,producto "
                        Else
                            Isql = "select sector as local, tipoproducto, sum(cantidad) as venta from " & sNombre & " group by sector,tipoproducto "
                        End If
                Else ' ambos
                
                    If optTipo(0).value = True Then ' detallao
                        If optValor(2).value = True Then ' costo
                            Isql = "select sector+' Cant' as local , producto, sum(cantidad) as venta  from " & sNombre & " group by sector,producto  union select sector+' Mont' as local, producto, sum(costo) as venta  from " & sNombre & " group by sector,producto "
                        Else ' neto /venta
                            Isql = "select sector+' Cant'  as local, producto, sum(cantidad) as venta  from " & sNombre & " group by sector,producto  union select sector+' Mont' as local, producto, sum(venta) as venta  from " & sNombre & " group by sector,producto "
                        End If
                        
                    Else ' resumido
                        If optValor(2).value = True Then ' costo
                            Isql = "select sector+' Cant'  as local, tipoproducto, sum(cantidad) as venta  from " & sNombre & " group by sector,tipoproducto  union select sector+' Mont' as local, tipoproducto, sum(costo) as venta  from " & sNombre & " group by sector,tipoproducto "
                        Else ' neto /venta
                            Isql = "select sector+' Cant' as local, tipoproducto, sum(cantidad) as venta  from " & sNombre & " group by sector,tipoproducto  union select sector+' Mont' as local,  tipoproducto, sum(venta) as venta  from " & sNombre & " group by sector,tipoproducto "
                        End If
                        
                    End If
                End If
                
               Set rsReporte = Lib.OpenRecordset(Isql, Cn)
                Resumidol.DiscardSavedData
                Resumidol.Database.SetDataSource rsReporte
                Resumidol.PaperOrientation = crLandscape
                
                Resumidol.ReportTitle = "Por Sectores. Consolidado entre: " & Me.dtpFecIni.value & "  y " & Me.dtpFecFin.value
                If Me.chkFHoraria.value = 1 Then
                    Resumidol.ReportTitle = Resumidol.ReportTitle & " Franja horaria: " & Format(Me.dtpHorIni.value, "HH:MM") & " - " & Format(Me.dtpHorFin.value, "HH:MM")
                End If
                Resumidol.Text20.SetText sRazonSocial
                Resumidol.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos") '& " (Monto Expresado en Soles)"
                Resumidol.Text6.SetText "Paloteo Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Valorización: " & IIf(optRanking(0).value = True, "Cantidad", IIf(optValor(0).value = True, optValor(0).Caption & " (Montos Expresados en Soles)", IIf(optValor(1).value = True, optValor(1).Caption & " (Montos Expresados en Soles)", optValor(2).Caption & " (Montos Expresados en Soles)")))
                frmEmite.CRViewer.ReportSource = Resumidol
            End If
End Sub
Private Sub configuraListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = devuelveServidoresConectados
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With
End Sub

Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
       End If
   ' validaCantidadSeleccion
    
    
End Sub
Public Function verificaSectorSeleccionado(sectors As String) As Boolean
    verificaSectorSeleccionado = True
    Dim q As Integer
        For q = 0 To 500
            If sector(q) = sectors Then
                verificaSectorSeleccionado = False
                Exit Function
            End If
        Next q
End Function
Public Sub limpiaSector()
    Dim q As Integer
        For q = 0 To 500
            sector(q) = ""
        Next q
End Sub


Private Sub obtieneListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = rsListaServidores
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With

End Sub
Public Sub validaCantidadSeleccion()
Dim p, w As Integer
Dim s As Integer
    p = 0
    s = 0
    limpiaSector
    For w = 0 To lstServidores.ListItems.Count - 1
        If lstServidores.ListItems.Item(w + 1).Checked = True Then
            p = p + 1
            If verificaSectorSeleccionado(lstServidores.ListItems(w + 1).SubItems(1)) Then
                sector(w) = lstServidores.ListItems(w + 1).SubItems(1)
            End If
        End If
    Next w
    For w = 0 To 500
        If sector(w) <> "" Then
            s = s + 1
        End If
    Next w
    If Me.optOpcion(1).value = True Then
        If p > 8 Then
            Me.cmdOpcion(0).Enabled = False
        Else
            Me.cmdOpcion(0).Enabled = True
        End If
    Else
        If s > 8 Then
            Me.cmdOpcion(0).Enabled = False
        Else
            Me.cmdOpcion(0).Enabled = True
        End If
    End If
End Sub


Private Sub optRanking_Click(Index As Integer)
    If Me.optRanking(0).value = True Then
        Frame4.Enabled = False
    Else
        Frame4.Enabled = True
    End If
End Sub

Private Sub optValor_Click(Index As Integer)
    Select Case Index
        Case 3
            Frame6.Enabled = False
        Case Else
            Frame6.Enabled = True
    End Select
End Sub

