VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepPaloteoPropiedades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Propiedades"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepPaloteoPropiedades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
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
      Left            =   4635
      Picture         =   "frmRepPaloteoPropiedades.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7770
      Visible         =   0   'False
      Width           =   1455
   End
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
      Height          =   7710
      Left            =   30
      TabIndex        =   38
      Top             =   15
      Width           =   7620
      Begin VB.CheckBox chkOrigenVenta 
         Caption         =   "Todos los Origen Venta"
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
         Left            =   5115
         TabIndex        =   61
         Top             =   645
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkCFacturados 
         Caption         =   "Valorizar con Cero la Pre Venta Facturada"
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   5320
         Width           =   3375
      End
      Begin VB.CheckBox chkUnidadNegocio 
         Caption         =   "Todas las Und.Negoc."
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
         TabIndex        =   9
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2265
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
         Left            =   5130
         TabIndex        =   52
         Top             =   5400
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   1116
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   55
            Top             =   1660
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   1388
            Width           =   2025
         End
      End
      Begin VB.CheckBox chkPropiedad 
         Caption         =   "Todos las Propiedades"
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
         TabIndex        =   21
         Top             =   3885
         Value           =   1  'Checked
         Width           =   2295
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
         Index           =   2
         Left            =   4230
         Picture         =   "frmRepPaloteoPropiedades.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3840
         Width           =   765
      End
      Begin VB.TextBox txtPropiedad 
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3870
         Width           =   2265
      End
      Begin VB.CheckBox chkOperador 
         Caption         =   "Todos los Operadores"
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
         TabIndex        =   18
         Top             =   3525
         Value           =   1  'Checked
         Width           =   2250
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
         Left            =   5130
         TabIndex        =   11
         Top             =   2460
         Value           =   1  'Checked
         Width           =   1875
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
         TabIndex        =   13
         Top             =   2820
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkCaja 
         Caption         =   "Todas las Cajas"
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
         Top             =   1740
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Todas las Areas"
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
         TabIndex        =   5
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4230
         Width           =   2265
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3165
         Width           =   2265
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
         Left            =   4230
         Picture         =   "frmRepPaloteoPropiedades.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4215
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
         Index           =   0
         Left            =   4230
         Picture         =   "frmRepPaloteoPropiedades.frx":0A32
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3135
         Width           =   765
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
         Height          =   2040
         Left            =   1860
         TabIndex        =   45
         Top             =   5640
         Width           =   3135
         Begin VB.OptionButton optOpcion 
            Caption         =   "Propiedades"
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
            Index           =   4
            Left            =   165
            TabIndex        =   33
            Top             =   1530
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Area de Producción"
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
            Left            =   165
            TabIndex        =   31
            Top             =   750
            Width           =   2265
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Grupo / SubGrupo"
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
            Left            =   165
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Pedidos"
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
            Index           =   3
            Left            =   165
            TabIndex        =   32
            Top             =   1140
            Width           =   1755
         End
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
         Left            =   5130
         TabIndex        =   3
         Top             =   1035
         Value           =   1  'Checked
         Width           =   2310
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Productos"
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
         TabIndex        =   16
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2115
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
         Left            =   5130
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
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
         TabIndex        =   24
         Top             =   4245
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1860
         TabIndex        =   27
         Top             =   4980
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
         Format          =   56098817
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1860
         TabIndex        =   25
         Top             =   4590
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
         Format          =   56098817
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   5130
         TabIndex        =   26
         Top             =   4575
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   56098819
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5130
         TabIndex        =   28
         Top             =   4965
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   56098819
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   300
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
      Begin MSDataListLib.DataCombo cboMozo 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   1020
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   1380
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
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   1725
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
         Left            =   1860
         TabIndex        =   10
         Top             =   2445
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
         Left            =   1860
         TabIndex        =   12
         Top             =   2805
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
      Begin MSDataListLib.DataCombo cboOperador 
         Height          =   315
         Left            =   1860
         TabIndex        =   17
         Top             =   3510
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
      Begin MSDataListLib.DataCombo cboUnidadNegocio 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   2085
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
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   1860
         TabIndex        =   62
         Top             =   660
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origen de Venta :"
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
         Index           =   15
         Left            =   250
         TabIndex        =   63
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Negocio :"
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
         Index           =   8
         Left            =   300
         TabIndex        =   60
         Top             =   2145
         Width           =   1500
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Propiedad :"
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
         Left            =   810
         TabIndex        =   51
         Top             =   3930
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador :"
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
         Index           =   0
         Left            =   885
         TabIndex        =   50
         Top             =   3570
         Width           =   915
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
         Left            =   1155
         TabIndex        =   49
         Top             =   2505
         Width           =   645
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
         Left            =   825
         TabIndex        =   48
         Top             =   2865
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caja :"
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
         Index           =   11
         Left            =   1290
         TabIndex        =   47
         Top             =   1785
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Producción :"
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
         Index           =   3
         Left            =   255
         TabIndex        =   46
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meseros :"
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
         Index           =   9
         Left            =   960
         TabIndex        =   44
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
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
         Index           =   7
         Left            =   900
         TabIndex        =   43
         Top             =   3225
         Width           =   900
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
         Left            =   375
         TabIndex        =   42
         Top             =   360
         Width           =   1425
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
         Left            =   675
         TabIndex        =   41
         Top             =   5055
         Width           =   1125
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
         Left            =   570
         TabIndex        =   40
         Top             =   4665
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         Index           =   2
         Left            =   1170
         TabIndex        =   39
         Top             =   4290
         Width           =   630
      End
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
      Left            =   6090
      Picture         =   "frmRepPaloteoPropiedades.frx":0B34
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7770
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
      Left            =   3180
      Picture         =   "frmRepPaloteoPropiedades.frx":0C26
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7770
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
      Left            =   1725
      Picture         =   "frmRepPaloteoPropiedades.frx":1158
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7770
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   840
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepPaloteoPropiedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset

Dim RsTipoPedido As Recordset
Dim RsMozo As Recordset
Dim RsArea As Recordset
Dim RsCaja As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsProducto As Recordset
Dim RsOperador As Recordset
Dim RsPropiedad As Recordset
Dim Rsturno As Recordset
Dim RsUnidadNegocio As Recordset
Dim RsOrigen As Recordset

Dim sCriterio   As String
Dim xCriterio   As String
Dim sProducto   As String
Dim sTurno      As String
Dim sPropiedad  As String
Dim sTitulo     As String
Dim sTexto      As String

Dim Grupo     As New dsrPaloteoPropiedadG
Dim Area      As New dsrPaloteoPropiedadA
Dim Pedido    As New dsrPaloteoPropiedadP
Dim Detallado As New dsrPaloteoPropiedadD

Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With Me.cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta where Activo=1"
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboMozo
         Isql = "Select * from vMozo where substring(Codigo,1,1)<>'*'"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboArea
         Isql = "Select * from vArea"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboCaja
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from tCaja"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
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
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboOperador
         Isql = "Select * from vOperador"
         Set RsOperador = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOperador
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboUnidadNegocio
         Isql = "Select * from vUnidadNegocio"
         Set RsUnidadNegocio = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUnidadNegocio
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


Private Sub cboTipoPedido_Change()
    Me.cboOrigenVenta.BoundText = ""
    With cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta  " & IIf(cboTipoPedido.BoundText = "", "", "where CodCanalVenta = '" & cboTipoPedido.BoundText & "' and Activo=1") & " order by Descripcion "
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
    End With
End Sub
Private Sub chkArea_Click()
   If chkArea.Value = 1 Then
      cboArea.Enabled = False
      cboArea.Text = ""
   Else
      cboArea.Enabled = True
   End If
End Sub

Private Sub ChkCaja_Click()
   If chkCaja.Value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

Private Sub chkOrigenVenta_Click()
    If chkOrigenVenta.Value = 1 Then
        Me.cboOrigenVenta.Enabled = False
        Me.cboOrigenVenta.Text = ""
    Else
        Me.cboOrigenVenta.Enabled = True
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

Private Sub chkMozo_Click()
   If chkMozo.Value = 1 Then
      cboMozo.Enabled = False
      cboMozo.Text = ""
   Else
      cboMozo.Enabled = True
   End If
End Sub

Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
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

Private Sub chkPropiedad_Click()
   If chkPropiedad.Value = 1 Then
      sPropiedad = ""
      txtPropiedad.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
   End If
End Sub

Private Sub chkOperador_Click()
   If chkOperador.Value = 1 Then
      cboOperador.Enabled = False
      cboOperador.Text = ""
   Else
      cboOperador.Enabled = True
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

Private Sub chkUnidadNegocio_Click()
   If chkUnidadNegocio.Value = 1 Then
      cboUnidadNegocio.Enabled = False
      cboUnidadNegocio.Text = ""
   Else
      cboUnidadNegocio.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""
   sTitulo = ""
   
   If chkTurno.Value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
      
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Rsturno.RecordCount > 0 Then
         sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
         sCriterio = " MPEDIDO.tTurno ='" & sTurno & "'"
         xCriterio = " MDOCUMENTO.tTurno='" & sTurno & "'"
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
   Else
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
      xCriterio = " MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
   End If
      
   If chkOperador.Value = 0 Then
      If cboOperador.Text = "" Then
         MsgBox "Debe escoger un Operador", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vOperador.Codigo ='" & cboOperador.BoundText & "'"
   End If
      
   If chkUnidadNegocio.Value = 0 Then
      If cboUnidadNegocio.Text = "" Then
         MsgBox "Debe escoger una unidad de negocio", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and DPEDIDO.tUnidadNegocio ='" & cboUnidadNegocio.BoundText & "'"
   End If
      
   If chkMozo.Value = 0 Then
      If cboMozo.Text = "" Then
         MsgBox "Debe escoger un Mesero", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tMozo ='" & cboMozo.BoundText & "'"
   End If
   
   If chkTipoPedido.Value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.BoundText & "'"
   End If
    If Me.chkOrigenVenta.Value = 0 Then
      If Me.cboOrigenVenta.Text = "" Then
         MsgBox "Debe escoger un Origen Venta", vbCritical, sMensaje
         Exit Sub
      End If
    End If
   If chkArea.Value = 0 Then
      If cboArea.Text = "" Then
         MsgBox "Debe escoger un Area", vbCritical, sMensaje
        Exit Sub
      End If
      sCriterio = sCriterio & " and Area ='" & cboArea.Text & "'"
   End If
      
   If chkCaja.Value = 0 Then
      If cboCaja.Text = "" Then
         MsgBox "Debe escoger una Caja", vbCritical, sMensaje
        Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tCaja ='" & cboCaja.BoundText & "'"
   End If
      
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.Grupo='" & cboGrupo.Text & "'"
   End If
   
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and subGrupo='" & cboSubGrupo.Text & "'"
   End If
   
   If chkProducto.Value = 0 Then
      If sProducto = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and DPEDIDO.tCodigoProducto ='" & sProducto & "'"
   End If
   
   If chkPropiedad.Value = 0 Then
      If sPropiedad = "" Then
         MsgBox "Debe escoger una Propiedad", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and TPRODUCTOPROPIEDAD.tCodigoPropiedad ='" & sPropiedad & "'"
   End If
   
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

               If optOpcion(1).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = True
                  Grupo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Grupo.PaperOrientation = crPortrait
               ElseIf optOpcion(2).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Area.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Area.PaperOrientation = crPortrait
               ElseIf optOpcion(3).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Pedido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Pedido.PaperOrientation = crPortrait
               Else
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
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
               If optOpcion(1).Value = True Then
                  Grupo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Grupo.PaperOrientation = crPortrait
                  Grupo.PrintOut
               ElseIf optOpcion(2).Value = True Then
                  Area.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Area.PaperOrientation = crPortrait
                  Area.PrintOut
               ElseIf optOpcion(3).Value = True Then
                  Pedido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Pedido.PaperOrientation = crPortrait
                  Pedido.PrintOut
               Else
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
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
               If optOpcion(1).Value = True Then
                  Grupo.ExportOptions.FormatType = 21
                  Grupo.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Grupo.ExportOptions.DiskFileName = cmdSave.FileName
                  Grupo.Export False
               ElseIf optOpcion(2).Value = True Then
                  Area.ExportOptions.FormatType = 21
                  Area.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Area.ExportOptions.DiskFileName = cmdSave.FileName
                  Area.Export False
               ElseIf optOpcion(3).Value = True Then
                  Pedido.ExportOptions.FormatType = 21
                  Pedido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Pedido.ExportOptions.DiskFileName = cmdSave.FileName
                  Pedido.Export False
               Else
                  Detallado.ExportOptions.FormatType = 21
                  Detallado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado.Export False
               End If
   End Select
End Sub

Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   Select Case Index
          Case Is = 0 ' producto
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
          
          Case Is = 1 ' turno
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
          
          Case Is = 2 ' propiedad
'               Isql = "Select tCodigoCliente, tIdentidad as Codigo, tApecom as Descripcion from TCOMPANIA order by tApeCom"
'               frmBusca.cboCriterio.Enabled = True
'               frmBusca.nPredeterm = 2
'               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "tCodigoCliente", 1200, 0, 0, "", _
'                                                      "Identificador", 2, "Codigo", 1500, 0, 0, "", _
'                                                      "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
'               frmBusca.Show vbModal
'               If Not wEnter Then
'                  Exit Sub
'               End If
'               sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
               Isql = "select codigo, descripcion,nprecio from vpropiedad where toperador like '" & cboOperador.BoundText & "%' order by 2"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "codigo", 1200, 0, 0, "", _
                                                      "Propiedad", 2, "descripcion", 4000, 0, 0, "", _
                                                      "Precio", 2, "nprecio", 1200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sPropiedad = sCodigo
               Me.txtPropiedad.Text = sDescrip
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
   Else
      cmdBusca(1).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
 
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    
    cboTipoPedido.Enabled = False
    cboTipoPedido.Text = ""
    cboMozo.Enabled = False
    cboMozo.Text = ""
    cboArea.Enabled = False
    cboArea.Text = ""
    cboCaja.Enabled = False
    cboCaja.Text = ""
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    cboOperador.Enabled = False
    cboOperador.Text = ""
    cboUnidadNegocio.Enabled = False
    cboUnidadNegocio.Text = ""
    Me.cboOrigenVenta.Enabled = False
    Me.cboOrigenVenta.Text = ""
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    cmdBusca(2).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsTipoPedido = Nothing
   Set RsMozo = Nothing
   Set RsCaja = Nothing
   Set RsGrupo = Nothing
   Set RsSubgrupo = Nothing
   Set RsProducto = Nothing
   Set RsOperador = Nothing
   Set RsUnidadNegocio = Nothing
   Set RsPropiedad = Nothing
   Set frmRepPaloteoPropiedades = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PaloteoPropiedad", Cn) Then
       Set oComando = Nothing
       Exit Sub
       Dim tipopedido As String
    End If
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    oComando.CreateParameter "@flagOrigenProduccion", adBoolean, adParamInput, 1, optOrigen(0).Value
    oComando.CreateParameter "@flagOrigenVenta", adBoolean, adParamInput, 1, optOrigen(1).Value
    oComando.CreateParameter "@flagOrigenCortesia", adBoolean, adParamInput, 1, optOrigen(2).Value
    oComando.CreateParameter "@flagOrigenCuentaCorriente", adBoolean, adParamInput, 1, optOrigen(3).Value
    oComando.CreateParameter "@flagOrigenCombinacion", adBoolean, adParamInput, 1, optOrigen(4).Value
    oComando.CreateParameter "@flagOrigenCargos", adBoolean, adParamInput, 1, optOrigen(5).Value
    oComando.CreateParameter "@flagOrigenPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).Value
    oComando.CreateParameter "@flagTurnoOFecha", adBoolean, adParamInput, 1, chkTurno.Value
    oComando.CreateParameter "@flagOpcion1", adBoolean, adParamInput, 1, optOpcion(1).Value
    oComando.CreateParameter "@flagOpcion2", adBoolean, adParamInput, 1, optOpcion(2).Value
    oComando.CreateParameter "@flagOpcion3", adBoolean, adParamInput, 1, optOpcion(3).Value
    oComando.CreateParameter "@flagOpcion4", adBoolean, adParamInput, 1, optOpcion(4).Value
    If BDLink = "1" Then
        oComando.CreateParameter "@dbAlmacen", adVarChar, adParamInput, 35, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
    Else
        oComando.CreateParameter "@dbAlmacen", adVarChar, adParamInput, 35, sAlmacenMDB   'CASO CORP.FERNANDEZ
    End If
    oComando.CreateParameter "@turno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@tOperador", adVarChar, adParamInput, 30, cboOperador.BoundText
    oComando.CreateParameter "@tUnidadNegocio", adVarChar, adParamInput, 30, cboUnidadNegocio.BoundText
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 30, cboMozo.BoundText
    If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.Value = 0 Then
        tipopedido = cboTipoPedido.BoundText
    Else
        tipopedido = ""
    End If
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tArea", adVarChar, adParamInput, 30, cboArea.Text
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 30, cboCaja.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@tcodigoPropiedad", adVarChar, adParamInput, 30, sPropiedad
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@lAlmacen", adBoolean, adParamInput, 1, lAlmacen
    oComando.CreateParameter "@OrigenVenta", adVarChar, adParamInput, 30, Me.cboOrigenVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
   If optOpcion(1).Value = True Then
      Grupo.DiscardSavedData
      Grupo.Text25.SetText localConectado
      Grupo.Database.SetDataSource rsReporte
      Grupo.Text10.SetText "Paloteo de Propiedades de " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))) & " Agrupado por Grupo/SubGrupo"
      Grupo.Text13.SetText sRazonSocial
      Grupo.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & " Expresado en " & sMonedaN
      frmEmite.CRViewer.ReportSource = Grupo
   End If
   
   If optOpcion(2).Value = True Then
      Area.DiscardSavedData
      Area.Text25.SetText localConectado
      Area.Database.SetDataSource rsReporte
      Area.Text10.SetText "Paloteo de Propiedades de " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))) & " Agrupado por Areas"
      Area.Text13.SetText sRazonSocial
      Area.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & " Expresado en " & sMonedaN
      frmEmite.CRViewer.ReportSource = Area
   End If
            
   If optOpcion(3).Value = True Then
      Pedido.DiscardSavedData
      Pedido.Text25.SetText localConectado
      Pedido.Database.SetDataSource rsReporte
      Pedido.Text10.SetText "Paloteo de Propiedades " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))) & " por Pedidos"
      Pedido.Text13.SetText sRazonSocial
      Pedido.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      frmEmite.CRViewer.ReportSource = Pedido
   End If
   
   If optOpcion(4).Value = True Then
      Detallado.DiscardSavedData
      Detallado.Text25.SetText localConectado
      Detallado.Database.SetDataSource rsReporte
      Detallado.Text10.SetText "Paloteo de Propiedades " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", "Combinaciones"))))) & " por Propiedades"
      Detallado.Text13.SetText sRazonSocial
      Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      frmEmite.CRViewer.ReportSource = Detallado
   End If

End Sub


