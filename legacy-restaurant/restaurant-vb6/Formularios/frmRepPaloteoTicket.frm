VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepPaloteoTicket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Producción por Ticketera"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepPaloteoTicket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7590
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
      Height          =   7230
      Left            =   30
      TabIndex        =   29
      Top             =   0
      Width           =   7545
      Begin VB.Frame Frame2 
         Caption         =   "Opciones de Impresión"
         Height          =   975
         Left            =   240
         TabIndex        =   55
         Top             =   6120
         Width           =   4695
         Begin VB.OptionButton optOrden 
            Caption         =   "Cod. Producto"
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
            Left            =   3000
            TabIndex        =   58
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optOrden 
            Caption         =   "Alfabetico"
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
            Left            =   1080
            TabIndex        =   57
            Top             =   600
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkSubtotales 
            Caption         =   "Mostrar Total por Producto"
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
            Left            =   1080
            TabIndex        =   56
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mostrar :"
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
            Index           =   12
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orden :"
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
            Left            =   240
            TabIndex        =   59
            Top             =   600
            Width           =   645
         End
      End
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
         TabIndex        =   52
         Top             =   1380
         Value           =   1  'Checked
         Width           =   2505
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
         TabIndex        =   44
         Top             =   4860
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   1116
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   47
            Top             =   1660
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   1388
            Width           =   2025
         End
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Todos los Clientes"
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
         TabIndex        =   19
         Top             =   3630
         Value           =   1  'Checked
         Width           =   2115
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
         Picture         =   "frmRepPaloteoTicket.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3585
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
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
         TabIndex        =   17
         Top             =   3615
         Width           =   2265
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
         TabIndex        =   9
         Top             =   2100
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
         TabIndex        =   20
         Top             =   3990
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
         Top             =   3225
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
         Picture         =   "frmRepPaloteoTicket.frx":010E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3960
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
         Picture         =   "frmRepPaloteoTicket.frx":0210
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3195
         Width           =   765
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
         TabIndex        =   7
         Top             =   1725
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
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2115
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
         Top             =   2865
         Value           =   1  'Checked
         Width           =   2175
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
         TabIndex        =   5
         Top             =   975
         Value           =   1  'Checked
         Width           =   1905
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
         Top             =   2475
         Value           =   1  'Checked
         Width           =   1875
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
         TabIndex        =   22
         Top             =   4005
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkSalon 
         Caption         =   "Todos los Salones"
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
         Top             =   210
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   5130
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1860
         TabIndex        =   25
         Top             =   4785
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   56426497
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1860
         TabIndex        =   23
         Top             =   4365
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   56426497
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3570
         TabIndex        =   24
         Top             =   4350
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
         Format          =   56426499
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3570
         TabIndex        =   26
         Top             =   4770
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
         Format          =   56426499
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   585
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
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   195
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
         Left            =   1860
         TabIndex        =   4
         Top             =   960
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
         TabIndex        =   6
         Top             =   1710
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
         Top             =   2460
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
         Top             =   2850
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
      Begin MSDataListLib.DataCombo cboImpresora 
         Height          =   360
         Left            =   1845
         TabIndex        =   42
         Top             =   5220
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   1830
         TabIndex        =   53
         Top             =   1320
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
         Left            =   240
         TabIndex        =   54
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora :"
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
         Left            =   825
         TabIndex        =   43
         Top             =   5310
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Index           =   10
         Left            =   1065
         TabIndex        =   41
         Top             =   3675
         Width           =   720
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
         Left            =   240
         TabIndex        =   40
         Top             =   2145
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
         Left            =   1110
         TabIndex        =   39
         Top             =   1770
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
         Left            =   885
         TabIndex        =   38
         Top             =   3285
         Width           =   900
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
         Left            =   810
         TabIndex        =   37
         Top             =   2910
         Width           =   975
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
         Left            =   360
         TabIndex        =   36
         Top             =   1020
         Width           =   1425
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
         Left            =   1140
         TabIndex        =   35
         Top             =   2520
         Width           =   645
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
         Left            =   660
         TabIndex        =   34
         Top             =   4860
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
         Left            =   555
         TabIndex        =   33
         Top             =   4440
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
         Left            =   1155
         TabIndex        =   32
         Top             =   4050
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salón :"
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
         Left            =   1170
         TabIndex        =   31
         Top             =   255
         Width           =   615
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
         Left            =   180
         TabIndex        =   30
         Top             =   645
         Width           =   1605
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
      Index           =   0
      Left            =   6090
      Picture         =   "frmRepPaloteoTicket.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Ticketera"
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
      Left            =   4620
      Picture         =   "frmRepPaloteoTicket.frx":0404
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   90
      Top             =   7305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepPaloteoTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim RsReporte As Recordset
Dim RsTempoE As Recordset
Dim RsTempoT As Recordset
Dim RsTempoC As Recordset
Dim RsTempoP As Recordset
Dim RsTurno As Recordset
Dim RsSalon As Recordset
Dim RsMozo As Recordset
Dim RsTipoPedido As Recordset
Dim RsTipo As Recordset
Dim RsArea As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsProducto As Recordset
Dim RsCliente As Recordset
Dim RsImpresora As Recordset
Dim RsOrigen As Recordset

Dim sCriterio As String
Dim sProducto As String
Public sTurno As String
Dim sCliente As String
Dim sFiltro As String
Dim sTitulo As String
Dim xCriterio As String
Dim sLinea As String

Sub LlenaCombos()
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With

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
    
    With Me.cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta where Activo=1"
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboMozo
         Isql = "Select * from vMozo"
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
    
    With cboImpresora
         Isql = "Select * from TIMPRESORA Where tCaja = '" & sCaja & "'"
         Set RsImpresora = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsImpresora
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
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
   If chkArea.value = 1 Then
      cboArea.Enabled = False
      cboArea.Text = ""
   Else
      cboArea.Enabled = True
   End If
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub
Private Sub chkOrigenVenta_Click()
    If chkOrigenVenta.value = 1 Then
        Me.cboOrigenVenta.Enabled = False
        Me.cboOrigenVenta.Text = ""
    Else
        Me.cboOrigenVenta.Enabled = True
    End If
End Sub
Private Sub chkMozo_Click()
   If chkMozo.value = 1 Then
      cboMozo.Enabled = False
      cboMozo.Text = ""
   Else
      cboMozo.Enabled = True
   End If
End Sub

Private Sub chkGrupo_Click()
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

Private Sub chkSalon_Click()
   If chkSalon.value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
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

Private Sub chkproducto_Click()
   If chkProducto.value = 1 Then
      sProducto = ""
      txtProducto.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkCliente_Click()
   If chkCliente.value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
   End If
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""
   sTitulo = ""
   sFiltro = ""
   
   If cboImpresora.Text = "" Then
      MsgBox "Escoja una impresora", vbCritical, sMensaje
      Exit Sub
   End If
   
   If chkTurno.value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
      
      Set RsTurno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If RsTurno.RecordCount > 0 Then
         sTitulo = " Del " & Format(RsTurno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(RsTurno!fFinal), Format(RsTurno!finicial, "dd/mmm/yyyy") & " 23:59", RsTurno!fFinal), "dd/mmm/yyyy HH:mm")
         sCriterio = " MPEDIDO.tTurno ='" & sTurno & "'"
         xCriterio = " MDOCUMENTO.tTurno='" & sTurno & "'"
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
   Else
      If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "
      xCriterio = " MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "
   End If
  
   If chkSalon.value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salón", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & "and MPEDIDO.tSalon ='" & cboSalon.BoundText & "'"
   End If
   
   If chkTipo.value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de Producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tTipoProducto ='" & cboTipoProducto.BoundText & "'"
   End If
   
   If chkMozo.value = 0 Then
      If cboMozo.Text = "" Then
         MsgBox "Debe escoger un " & Mesero, vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tMozo ='" & cboMozo.BoundText & "'"
   End If
   
   If chkTipoPedido.value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.BoundText & "'"
   End If
   
    If Me.chkOrigenVenta.value = 0 Then
      If Me.cboOrigenVenta.Text = "" Then
         MsgBox "Debe escoger un Origen Venta", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.CodigoOrigenVenta ='" & Me.cboOrigenVenta.BoundText & "'"
   End If
   If chkArea.value = 0 Then
      If cboArea.Text = "" Then
         MsgBox "Debe escoger un Area", vbCritical, sMensaje
        Exit Sub
      End If
      
      sCriterio = sCriterio & " and Area ='" & cboArea.Text & "'"
   End If
      
   If chkGrupo.value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.Grupo='" & cboGrupo.Text & "'"
   End If
   
   If chkSubGrupo.value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and subGrupo='" & cboSubGrupo.Text & "'"
   End If
   
   If chkProducto.value = 0 Then
      If sProducto = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and DPEDIDO.tCodigoProducto ='" & sProducto & "'"
   End If
   
   If chkCliente.value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger un Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sFiltro = " and MDOCUMENTO.tCodigoCliente ='" & sCliente & "'"
   End If
   
   Genera
   If RsReporte.EOF = True Then
      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
               
   'Configura la impresora la impresion Font
   Imprimir (cboImpresora.BoundText)
   Printer.FontName = sFont
   Printer.FontBold = False
   
   'Cabecera
   ImprimeXCentro "Paloteo de " & IIf(optOrigen(0).value = True, "Produccion", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesias", IIf(optOrigen(3).value = True, "Cuentas Corrientes", "Combinaciones")))), 40
   ImprimeXCentro sRazonSocial, 40
   Printer.Print ""
   ImprimeXCentro IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo, 40
   Printer.Print ""
   Printer.Print "Fecha : " & FechaServidor()
    Printer.Print String(40, "-")
   
   
   'CANAL DE VENTA GCAA_29032022
    If cboTipoPedido.Text <> "" Then
        Printer.Print "Tipo de Venta: " & UCase(cboTipoPedido.Text)
        Printer.Print " "
    End If
   Printer.Print String(40, "-")
   
   '11092018 AGREGAR TOTAL FORMATO PARA LA IMPRESION EN TICKET
   If (chkSubtotales.value = False) Then
        Printer.Print "Producto                        Cantidad"
   Else
        Printer.Print Left("Producto" & String(20, " "), 20) & " " & Right(String(8, " ") & "Cantidad", 8) & " " & Right(String(10, " ") & "Total", 10)
   End If
   Printer.Print String(40, "-")
  
    
    
    
   'Detalle
   RsReporte.MoveFirst
   Dim nTotal As Double
   Dim nTotalCant As Double '11092018 AGREGAR TOTAL FORMATO PARA LA IMPRESION EN TICKET
   Dim xParte1 As String
   Dim xParte2 As String
   nTotal = 0
   
   Do While Not RsReporte.EOF
      nTotal = nTotal + RsReporte!Venta
    
    '11092018 AGREGAR TOTAL FORMATO PARA LA IMPRESION EN TICKET
    nTotalCant = nTotalCant + RsReporte!Cantidad
    If (chkSubtotales.value = False) Then
        xParte1 = Left(RsReporte!Producto & String(28, " "), 28)
        xParte2 = Right(String(11, " ") & Format(RsReporte!Cantidad, "####,##0.00"), 11)
    Else
        xParte1 = Left(RsReporte!Producto & String(20, " "), 20)
        xParte2 = Right(String(8, " ") & Format(RsReporte!Cantidad, "#,##0.00"), 8)
        xParte2 = xParte2 & " " & Right(String(10, " ") & Format(RsReporte!Venta, "###,##0.00"), 10)
    End If
      sLinea = xParte1 & " " & xParte2
      '-----------------------------------------------------
      If pais = "003" Then
        sLinea = Replace(Replace(Replace(sLinea, ",", "|"), ".", ","), "|", ".")
      End If
      '------------------------------------------------------
      Printer.Print sLinea
      RsReporte.MoveNext
   Loop
   Printer.Print ""
   Printer.Print String(40, "-")
   
   '11092018 AGREGAR TOTAL FORMATO PARA LA IMPRESION EN TICKET
   If (chkSubtotales.value = False) Then
        xParte1 = Left("            Total : " & String(28, " "), 28)
        xParte2 = Right(String(11, " ") & Format(nTotal, "####,##0.00"), 11)
   Else
        xParte1 = Left("            Total : " & String(20, " "), 20)
        xParte2 = Right(String(8, " ") & Format(nTotalCant, "####,##0.00"), 8)
        xParte2 = xParte2 & " " & Right(String(10, " ") & Format(nTotal, "####,##0.00"), 10)
   End If
   
   sLinea = xParte1 & " " & xParte2
   '----------------------------------------------------------------------
   If pais = "003" Then
    sLinea = Replace(Replace(Replace(sLinea, ",", "|"), ".", ","), "|", ".")
   End If
   '---------------------------------------------------------------------
   Printer.Print sLinea ' xParte1 & " " & xParte2
   Printer.Print ""
   Printer.EndDoc
   Screen.MousePointer = vbDefault
               
   Genera
   If RsReporte.EOF = True Then
      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
     
   Screen.MousePointer = vbDefault
   If sModulo = "INFOREST" Then
      Unload Me
   End If
End Sub


Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   Select Case Index
          Case Is = 0
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
          
          Case Is = 1
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
          
          Case Is = 2
            '   Isql = "Select tCodigoCliente, tIdentidad as Codigo, tApecom as Descripcion from TCOMPANIA order by tApeCom"
               Isql = "Select codigo tcodigoCliente, Identidad as Codigo, tnomsoc as Descripcion from vCOMPANIA order by tnomsoc"
               frmBusca.cboCriterio.Enabled = True
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "tCodigoCliente", 1200, 0, 0, "", _
                                                      "Identificador", 2, "Codigo", 1500, 0, 0, "", _
                                                      "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
               txtCliente.Text = sDescrip
   
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub chkTurno_Click()
   If chkTurno.value = 1 Then
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
    sNombre = ""
    
    'JCDPFH 11092018 SE AGREGO EL COD producto para poder realizar sort
    'sNombre = dbTemporal(sCaja, 15,"tLocal", "nVarChar(3)", "Local", "nVarChar(30)", "Salon", "nVarChar(130)", "tMesa", "nVarChar(3)", "TipoProducto", "nVarChar(130)", "Grupo", "nVarChar(150)", "SubGrupo", "nVarChar(150)", "Producto", "nVarChar(150)", "Cantidad", "Float", "Venta", "Float", "Pedido", "nVarchar(10)", "Documento", "nVarChar(20)", "Fecha", "DateTime", "tTipoPedido", "nVarchar(2)", "Area", "nVarChar(30)")
    sNombre = dbTemporal(sCaja, 16, "tCodProducto", "nVarChar(8)", "tLocal", "nVarChar(3)", "Local", "nVarChar(30)", "Salon", "nVarChar(130)", "tMesa", "nVarChar(3)", "TipoProducto", "nVarChar(130)", "Grupo", "nVarChar(150)", "SubGrupo", "nVarChar(150)", "Producto", "nVarChar(150)", "Cantidad", "Float", "Venta", "Float", "Pedido", "nVarchar(10)", "Documento", "nVarChar(20)", "Fecha", "DateTime", "tTipoPedido", "nVarchar(2)", "Area", "nVarChar(30)")
    dtpFecIni.value = FechaServidor()
    dtpFecFin.value = FechaServidor()
    cboTipoProducto.Enabled = False
    cboTipoProducto.Text = ""
    cboSalon.Enabled = False
    cboTipoPedido.Enabled = False
    cboSalon.Text = ""
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    cboMozo.Enabled = False
    cboMozo.Text = ""
    cboArea.Enabled = False
    cboArea.Text = ""
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    cmdBusca(2).Enabled = False
    Me.cboOrigenVenta.Enabled = False
    Me.cboOrigenVenta.Text = ""
    cboImpresora.BoundText = sPreCuenta
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsReporte = Nothing
   Set RsTempoE = Nothing
   Set RsTempoT = Nothing
   Set RsTempoC = Nothing
   Set RsTempoP = Nothing
   Set RsTurno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsGrupo = Nothing
   Set RsMozo = Nothing
   Set RsTipo = Nothing
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   Cn.Execute "delete from " & sNombre
   
   'JCDPFH 11092018 SE AGREGO EL tCodProducto para poder realizar sort -----
   'into " & sNombre & " (tLocal,   ===> into " & sNombre & " (tCodProducto, tLocal,
   'SELECT dbo.vSalon.tLocal,       ===> SELECT dbo.vProducto.Codigo,dbo.vSalon.tLocal,
   
   
   'Produccion
   If optOrigen(0).value = True Then
   
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo,dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM  dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa " & _
             "where MPEDIDO.tEstadoPedido <> '03' and DPEDIDO.tEstadoItem='N' and " & sCriterio
   End If
   
   'Venta
   If optOrigen(1).value = True Then
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.MDOCUMENTO RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON dbo.MDOCUMENTO.tDocumento = dbo.DPEDIDO.tDocumento " & _
             "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sCriterio & " and " & xCriterio & sFiltro
   End If
   
   'Cortesia
   If optOrigen(2).value = True Then
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.TMESA RIGHT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa LEFT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPEDIDO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND dbo.DPEDIDO.tFacturado = 'C' and " & sCriterio
   End If
   
   'Cuentas Corrientes
   If optOrigen(3).value = True Then
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.TMESA RIGHT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.TMESA " & _
             "where isnull(MPEDIDO.tClienteCtaCte,'')<>'' and DPEDIDO.tEstadoItem='N' and " & sCriterio & " and dbo.DPEDIDO.tCodigoPedido not in " & _
             "(SELECT distinct dbo.DPEDIDO.tCodigoPedido FROM dbo.DPEDIDO RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPEDIDO.tDocumento = dbo.MDOCUMENTO.tDocumento Where ISNULL(tCodigoPedido,0)<>0 and " & xCriterio & ")"
   End If
   
   'Combinacion
   If optOrigen(4).value = True Then
      'Isql = "Insert into " & sNombre & " (tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.CPEDIDO.nCantidad, dbo.vProducto.nPrecioVenta AS Venta, dbo.CPEDIDO.tCodigoPedido AS Pedido, dbo.DPEDIDO.tDocumento AS Documento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.TMESA RIGHT OUTER JOIN dbo.MPEDIDO LEFT OUTER JOIN dbo.vSalon ON dbo.MPEDIDO.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.DPEDIDO RIGHT OUTER JOIN dbo.CPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem LEFT OUTER JOIN dbo.vProducto ON dbo.CPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.TMESA " & _
             "where MPEDIDO.tEstadoPedido <> '03' and " & sCriterio
        Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.CPEDIDO.nCantidad, dbo.vProducto.nPrecioVenta AS Venta, dbo.CPEDIDO.tCodigoPedido AS Pedido, dbo.DPEDIDO.tDocumento AS Documento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.TMESA RIGHT OUTER JOIN dbo.MPEDIDO LEFT OUTER JOIN dbo.vSalon ON dbo.MPEDIDO.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.DPEDIDO RIGHT OUTER JOIN dbo.CPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem LEFT OUTER JOIN dbo.vProducto ON dbo.CPEDIDO.tProductocombo = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.TMESA " & _
             "where MPEDIDO.tEstadoPedido <> '03' and " & sCriterio
             
   End If
   
   'Cargos
   If optOrigen(5).value = True Then
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM  dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa " & _
             "where MPEDIDO.tEstadoPedido = '05' and DPEDIDO.tEstadoItem='N' and " & sCriterio
   End If
   
   'Pedidos facturados
   If optOrigen(6).value = True Then
      Isql = "Insert into " & sNombre & " (tCodProducto, tLocal, Salon, tMesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, Pedido, Documento, Fecha, tTipoPedido, Area ) " & _
             "SELECT dbo.vProducto.Codigo, dbo.vSalon.tLocal, dbo.vSalon.Descripcion AS Salon, dbo.MPEDIDO.tMesa, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nVenta AS Venta, dbo.DPEDIDO.tCodigoPedido , dbo.DPEDIDO.tDocumento, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vProducto.Area " & _
             "FROM dbo.MDOCUMENTO RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON dbo.MDOCUMENTO.tDocumento = dbo.DPEDIDO.tDocumento " & _
             "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sCriterio & sFiltro
   End If
   
   Cn.Execute Isql
   Cn.Execute "update " & sNombre & " set Salon = 'Sin Salon' where isnull(Salon,'0') = '0'"
   Cn.Execute "update " & sNombre & " set Salon = '" & sBoton2 & "' where tTipoPedido ='02'"
   Cn.Execute "update " & sNombre & " set Salon = '" & sBoton3 & "' where tTipoPedido ='03'"
   Cn.Execute "update " & sNombre & " set Salon = '" & sBoton4 & "' where tTipoPedido ='04'"
   Cn.Execute "update " & sNombre & " set Salon = '" & sBoton5 & "' where tTipoPedido ='05'"
   Cn.Execute "update " & sNombre & " set Local = vLocal.Descripcion FROM " & sNombre & " INNER JOIN dbo.vLocal ON dbo." & sNombre & ".tLocal = dbo.vLocal.Codigo "
      
      
   Isql = "select tCodProducto,max(Local) as Local , max(Salon) as Salon, max(TipoProducto) as TipoProducto, Grupo, SubGrupo, Producto, sum(Cantidad) as Cantidad, sum(Venta) as Venta " & _
          "From " & sNombre & " group by tCodProducto,Grupo, SubGrupo, Producto order by Producto"
 
   Set RsReporte = Lib.OpenRecordset(Isql, Cn)
    If optOrden(1).value Then
        RsReporte.Sort = "tCodProducto ASC"
    End If
    
End Sub

