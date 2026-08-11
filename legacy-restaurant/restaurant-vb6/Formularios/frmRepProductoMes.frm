VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepProductoMes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Productos Por Meses"
   ClientHeight    =   8805
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepProductoMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
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
      Left            =   1283
      Picture         =   "frmRepProductoMes.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
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
      Left            =   2753
      Picture         =   "frmRepProductoMes.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
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
      Left            =   5663
      Picture         =   "frmRepProductoMes.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
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
      Height          =   8085
      Left            =   15
      TabIndex        =   42
      Top             =   0
      Width           =   8385
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
         Left            =   5760
         TabIndex        =   62
         Top             =   1250
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox chkCFacturados 
         Caption         =   "Valorizar con Cero la Pre Venta Facturada"
         Height          =   285
         Left            =   2400
         TabIndex        =   61
         Top             =   5450
         Width           =   4635
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
         Left            =   5730
         TabIndex        =   7
         Top             =   556
         Value           =   1  'Checked
         Width           =   1725
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
         Left            =   5730
         TabIndex        =   5
         Top             =   210
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
         Left            =   5730
         TabIndex        =   21
         Top             =   3345
         Value           =   1  'Checked
         Width           =   1830
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
         Left            =   5730
         TabIndex        =   9
         Top             =   902
         Value           =   1  'Checked
         Width           =   1890
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
         Left            =   5730
         TabIndex        =   23
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2145
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
         Left            =   5730
         TabIndex        =   25
         Top             =   4035
         Value           =   1  'Checked
         Width           =   2070
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
         Left            =   5730
         TabIndex        =   11
         Top             =   1605
         Value           =   1  'Checked
         Width           =   2250
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
         Left            =   5160
         TabIndex        =   46
         Top             =   5760
         Width           =   3045
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
            TabIndex        =   34
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(a) Ventas en el Rango"
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
            TabIndex        =   35
            Top             =   540
            Width           =   2295
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
            TabIndex        =   36
            Top             =   780
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
            TabIndex        =   37
            Top             =   1020
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   40
            Top             =   1740
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   41
            Top             =   1980
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Todos los Pedidos Facturad"
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
            TabIndex        =   38
            Top             =   1260
            Width           =   2790
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cancelados como Cortesias"
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
            Index           =   8
            Left            =   135
            TabIndex        =   39
            Top             =   1500
            Width           =   2745
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
         Left            =   4980
         Picture         =   "frmRepProductoMes.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3990
         Width           =   615
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   4020
         Width           =   2325
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
         Left            =   5730
         TabIndex        =   13
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1755
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   4365
         Width           =   2325
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
         Left            =   4980
         Picture         =   "frmRepProductoMes.frx":0C64
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4365
         Width           =   615
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
         Left            =   5730
         TabIndex        =   27
         Top             =   4380
         Value           =   1  'Checked
         Width           =   1965
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
         Left            =   5730
         TabIndex        =   15
         Top             =   2295
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.Frame Frame4 
         Caption         =   " Valorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   3150
         TabIndex        =   43
         Top             =   5760
         Width           =   1680
         Begin VB.OptionButton optValor 
            Caption         =   "Unidades"
            Height          =   240
            Index           =   3
            Left            =   135
            TabIndex        =   33
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   30
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   31
            Top             =   570
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   32
            Top             =   900
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Todas los Usuarios"
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
         Left            =   5730
         TabIndex        =   19
         Top             =   2985
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkUnidadNegocio 
         Caption         =   "Todas las Unidad Negocio"
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
         Left            =   5730
         TabIndex        =   17
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   540
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2655
         TabIndex        =   4
         Top             =   195
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   8
         Top             =   885
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1590
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   20
         Top             =   3330
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   22
         Top             =   3675
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1935
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   14
         Top             =   2280
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   18
         Top             =   2970
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   2640
         TabIndex        =   16
         Top             =   2625
         Width           =   2985
         _ExtentX        =   5265
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
      Begin MSComCtl2.DTPicker dtpAnual 
         Height          =   315
         Left            =   2640
         TabIndex        =   28
         Top             =   4755
         Width           =   2970
         _ExtentX        =   5239
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
         CustomFormat    =   "yyyy"
         Format          =   58720259
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   315
         Left            =   2640
         TabIndex        =   29
         Top             =   5160
         Width           =   2970
         _ExtentX        =   5239
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
         CustomFormat    =   "HH"
         Format          =   58720259
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   2640
         TabIndex        =   63
         Top             =   1245
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   1080
         TabIndex        =   64
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
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
         Index           =   13
         Left            =   2115
         TabIndex        =   60
         Top             =   4785
         Width           =   465
      End
      Begin VB.Label Label 
         Caption         =   "Hora de Cierre y Apertura  :"
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
         Left            =   225
         TabIndex        =   59
         Top             =   5190
         Width           =   2355
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
         Left            =   975
         TabIndex        =   58
         Top             =   600
         Width           =   1605
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
         Left            =   1965
         TabIndex        =   57
         Top             =   255
         Width           =   615
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
         Left            =   1935
         TabIndex        =   56
         Top             =   3390
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
         Left            =   1155
         TabIndex        =   55
         Top             =   945
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
         Left            =   1605
         TabIndex        =   54
         Top             =   3735
         Width           =   975
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
         Left            =   1680
         TabIndex        =   53
         Top             =   4080
         Width           =   900
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
         Left            =   1740
         TabIndex        =   52
         Top             =   1650
         Width           =   840
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
         Left            =   1035
         TabIndex        =   51
         Top             =   1995
         Width           =   1545
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
         Left            =   1860
         TabIndex        =   50
         Top             =   4425
         Width           =   720
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
         Left            =   2070
         TabIndex        =   49
         Top             =   2340
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
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
         Left            =   1800
         TabIndex        =   48
         Top             =   3030
         Width           =   780
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid. Negocio :"
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
         Left            =   1005
         TabIndex        =   47
         Top             =   2685
         Width           =   1575
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
      Left            =   4208
      Picture         =   "frmRepProductoMes.frx":0D66
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1335
      Top             =   7845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepProductoMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim rsReporteO   As Recordset
Dim RsTempoE    As Recordset
Dim RsTempoT    As Recordset
Dim RsTempoC    As Recordset
Dim RsTempoP    As Recordset
Dim Rsturno     As Recordset
Dim RsSalon     As Recordset
Dim RsMozo      As Recordset
Dim RsTipoPedido As Recordset
Dim RsTipo      As Recordset
Dim RsArea      As Recordset
Dim RsCaja      As Recordset
Dim RsGrupo     As Recordset
Dim RsSubgrupo  As Recordset
Dim RsProducto  As Recordset
Dim RsCliente   As Recordset
Dim RsUsuario   As Recordset
Dim RsUnidadNegocio  As Recordset
Dim RsOrigen As Recordset

Dim Resumido    As New dsrProductoXMes
Dim ResumidoMonto As New dsrProductoXMesMonto
Dim sCriterio   As String
Dim xCriterio   As String
Dim sProducto   As String
Dim sTurno      As String
Dim sCliente    As String
Dim sFiltro     As String
Dim sTitulo     As String
Dim sPrecio     As String
Dim sTexto      As String
Dim sTexto2     As String

Dim fInicio As Date
Dim fFinal As Date

Dim sDetalle As String

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
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from tCaja"  '?????????
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUsuario
         Isql = "Select tResumido as Codigo, tDetallado as Descripcion from tUsuario where lActivo = 1 and tGrupoUsuario<>'00'"
         Set RsUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUsuario
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

Private Sub chkUsuario_Click()
   If chkUsuario.Value = 1 Then
      cboUsuario.Enabled = False
      cboUsuario.Text = ""
   Else
      cboUsuario.Enabled = True
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

Private Sub chkSalon_Click()
   If chkSalon.Value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
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

Private Sub chkproducto_Click()
   If chkProducto.Value = 1 Then
      sProducto = ""
      txtProducto.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkCliente_Click()
   If chkCliente.Value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
       Unload Me
       Exit Sub
    End If
    sTitulo = ""
    sFiltro = ""
   
   ' armando fecha inicio fecha final +  hora de cierre
    Dim sFecha As String
    Dim sAno As String
    Dim sMes As String
    Dim sDia As String
    Dim j, i As Integer
    j = 0
    sFecha = ""
    sAno = LTrim(str(Year(dtpAnual.Value)))
   
    'sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & "01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & "12/31')"
    
    fInicio = DateAdd("H", Hour(dtpHora), Format("" & sAno & "/01/01", "yyyy/MM/DD")) ' & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = DateAdd("H", 24 + Hour(dtpHora), Format("" & sAno & "/12/31", "yyyy/MM/DD"))
    
    sTitulo = " Del " & Format(fInicio, "dd/mmm/yyyy HH:MM") & " Al " & Format(fFinal, "dd/mmm/yyyy HH:MM")
    
    If chkSalon.Value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salón", vbCritical, sMensaje
         Exit Sub
      End If
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
   
    If chkTipoPedido.Value = 0 Then
       If cboTipoPedido.Text = "" Then
          MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
          Exit Sub
       End If
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
    End If

      
    If chkCaja.Value = 0 Then
       If cboCaja.Text = "" Then
          MsgBox "Debe escoger una Caja", vbCritical, sMensaje
         Exit Sub
       End If
    End If
        
    If chkUsuario.Value = 0 Then
       If cboUsuario.Text = "" Then
          MsgBox "Debe escoger una Usuario", vbCritical, sMensaje
         Exit Sub
       End If
        sTitulo = sTitulo & " filtrado por el Usuario : " & cboUsuario.Text
    End If
    
    If chkUnidadNegocio.Value = 0 Then
       If cboUnidadNegocio.Text = "" Then
          MsgBox "Debe escoger una Unidad de negocio", vbCritical, sMensaje
         Exit Sub
       End If
       sTitulo = sTitulo & " filtrado por la Unidad de Negocio : " & cboUnidadNegocio.Text
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
   
   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger un Cliente", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporteO.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               If optValor(3).Value Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crLandscape
               Else
                  frmEmite.CRViewer.DisplayGroupTree = False
                  ResumidoMonto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ResumidoMonto.PaperOrientation = crLandscape
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               If rsReporteO.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               If optValor(3).Value Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crLandscape
                  Resumido.PrintOut
               
               Else
                  ResumidoMonto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ResumidoMonto.PaperOrientation = crLandscape
                  ResumidoMonto.PrintOut
               End If
                
          Case Is = 2 ' Salir
               Unload Me
          
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporteO.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optValor(3).Value Then
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
                Else
                  ResumidoMonto.ExportOptions.FormatType = 21
                  ResumidoMonto.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ResumidoMonto.ExportOptions.DiskFileName = cmdSave.FileName
                  ResumidoMonto.Export False
               End If
   End Select
   
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
'               Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
'               frmBusca.nPredeterm = 2
'               Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
'                                                      "Caja", 2, "tCaja", 1100, 0, 0, "", _
'                                                      "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
'                                                      "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
'
'               frmBusca.Show vbModal
'               If Not wEnter Then
'                  Exit Sub
'               End If
'               sTurno = sCodigo
'               txtTurno.Text = sCodigo
          
          Case Is = 2
                ' Isql = "select codigo as tCodigoCliente,tidentidad as Codigo, Descripcion from vcliente order by 3"
                ' Isql = "Select tCodigoCliente, tIdentidad as Codigo, tApecom as Descripcion from TCOMPANIA order by tApeCom"
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
'   If dtpFecIni.value > dtpFecFin.value Then
'      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
'      dtpFecFin.SetFocus
'   End If
End Sub



Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpAnual.Value = Date
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
    cboCaja.Enabled = False
    cboCaja.Text = ""
    cboUsuario.Enabled = False
    cboUnidadNegocio.Enabled = False
    cboUsuario.Text = ""
    cboUnidadNegocio.Text = ""
    Me.cboOrigenVenta.Enabled = False
    Me.cboOrigenVenta.Text = ""
    
    cmdBusca(0).Enabled = False
    cmdBusca(2).Enabled = False
    sDetalle = dbTemporal(sCaja, 43, "quiebre", "nvarchar(1)", "tGrupo", "nVarchar(200)", "tSubGrupo", "nVarchar(200)", "tcodigoProducto", "nVarchar(15)", "tProducto", "nVarchar(250)", "tipoProducto", "nVarchar(100)", "tAgrupacion", "nVarchar(100)", _
                                     "C1", "Float", "M1", "Float", _
                                     "C2", "Float", "M2", "Float", _
                                     "C3", "Float", "M3", "Float", _
                                     "C4", "Float", "M4", "Float", _
                                     "C5", "Float", "M5", "Float", _
                                     "C6", "Float", "M6", "Float", _
                                     "C7", "Float", "M7", "Float", _
                                     "C8", "Float", "M8", "Float", _
                                     "C9", "Float", "M9", "Float", _
                                     "C10", "Float", "M10", "Float", _
                                     "C11", "Float", "M11", "Float", _
                                     "C12", "Float", "M12", "Float", _
                                     "Enero", "Float", "Febrero", "Float", _
                                     "Marzo", "Float", "Abril", "Float", _
                                     "Mayo", "Float", "Junio", "Float", _
                                     "Julio", "Float", "Agosto", "Float", _
                                     "Setiembre", "Float", "Octubre", "Float", _
                                     "Noviembre", "Float", "Diciembre", "Float")
                                     
            '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set rsReporteO = Nothing
   Set RsTempoE = Nothing
   Set RsTempoT = Nothing
   Set RsTempoC = Nothing
   Set RsTempoP = Nothing
   Set Rsturno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsSubgrupo = Nothing
   Set RsCaja = Nothing
   Set RsMozo = Nothing
   Set RsTipo = Nothing
   Set RsUsuario = Nothing
   Set RsUnidadNegocio = Nothing
   Set frmRepPaloteo = Nothing
End Sub

Public Sub Genera()

   Dim sCostoPropiedad As String
   Dim sCombo As String
   Dim sCostoComboPropiedad As String
   Screen.MousePointer = vbHourglass
   If chkCFacturados.Value = 0 Then
               'Produccion
                If optOrigen(0).Value = True Then
                        If optValor(0).Value Then ' venta
                          sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Venta"
                          sTexto2 = "Total Venta"
                        ElseIf optValor(1).Value Then ' neto
                          sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Neto"
                          sTexto2 = "Total Neto"
                        ElseIf optValor(2).Value Then ' costo
                          sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Costo"
                          sTexto2 = "Total Costo"
                        Else ' cantidad
                          sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                          sTexto = "Unidades"
                          sTexto2 = "Total Unidades"
                        End If
                End If
                
            '   'Venta
               If optOrigen(1).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
            
               'Cortesia
               If optOrigen(2).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
            
               'Cuentas Corrientes
               If optOrigen(3).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
               
               'Combinacion
               If optOrigen(4).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
               
               'Cargos
               If optOrigen(5).Value = True Then
                        If optValor(0).Value Then
                          sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                          sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                          sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Costo"
                        Else
                          sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                          sTexto = "Unidades"
                          sTexto2 = "Total Unidades"
                        End If
               End If
            
               'Pedidos Facturados
               If optOrigen(6).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
                
                
            
               'Pedidos Facturados como cortesias
               If optOrigen(8).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
                End If
                
    Else 'c0n 0
               'Produccion
                If optOrigen(0).Value = True Then
                        If optValor(0).Value Then ' venta
                         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                        '  sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                          sTexto = "Precio Venta"
                          sTexto2 = "Total Venta"
                        ElseIf optValor(1).Value Then ' neto
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Precio Neto"
                          sTexto2 = "Total Neto"
                        ElseIf optValor(2).Value Then ' costo
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else  (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Precio Costo"
                          sTexto2 = "Total Costo"
                        Else ' cantidad
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Unidades"
                          sTexto2 = "Total Unidades"
                        End If
                End If
                
            '   'Venta
               If optOrigen(1).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
            
               'Cortesia
               If optOrigen(2).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
            
               'Cuentas Corrientes
               If optOrigen(3).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
               
               'Combinacion
               If optOrigen(4).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
            
               
               'Cargos
               If optOrigen(5).Value = True Then
                        If optValor(0).Value Then
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad end "
                          sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Precio Costo"
                        Else
                          sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                          sTexto = "Unidades"
                          sTexto2 = "Total Unidades"
                        End If
               End If
            
               'Pedidos Facturados
               If optOrigen(6).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
               End If
                
               'Pedidos Facturados como cortesias
               If optOrigen(8).Value = True Then
                        If optValor(0).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Venta"
                        ElseIf optValor(1).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Neto"
                        ElseIf optValor(2).Value Then
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Precio Costo"
                        Else
                           sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
                           sTexto = "Unidades"
                           sTexto2 = "Total Unidades"
                        End If
                End If
                
    End If
    
    Dim tipopedido As String
    If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.Value = 0 Then
        tipopedido = cboTipoPedido.BoundText
    Else
        tipopedido = ""
    End If
    
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PaloteoProduccionPorMes", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    


    oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).Value
    oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).Value
    oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).Value
    oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).Value
    oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).Value
    oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).Value
    oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).Value
    oComando.CreateParameter "@flagPedidosFacturadosCortesia", adBoolean, adParamInput, 1, optOrigen(8).Value
    oComando.CreateParameter "@flagPrecioCosto", adBoolean, adParamInput, 1, optValor(2).Value
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 30, cboSalon.BoundText
    oComando.CreateParameter "@tTipoProducto", adVarChar, adParamInput, 30, cboTipoProducto.BoundText
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 30, cboMozo.BoundText
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tArea", adVarChar, adParamInput, 30, cboArea.Text
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 30, cboCaja.BoundText
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 30, cboUsuario.BoundText
    oComando.CreateParameter "@tUnidadNegocio", adVarChar, adParamInput, 30, cboUnidadNegocio.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@tCodigoCliente", adVarChar, adParamInput, 30, sCliente
    oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 300, sPrecio
    oComando.CreateParameter "@sBoton2", adVarChar, adParamInput, 30, sBoton2
    oComando.CreateParameter "@sBoton3", adVarChar, adParamInput, 30, sBoton3
    oComando.CreateParameter "@sBoton4", adVarChar, adParamInput, 30, sBoton4
    oComando.CreateParameter "@sBoton5", adVarChar, adParamInput, 30, sBoton5
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
    oComando.CreateParameter "@flagvalorMuestra", adBoolean, adParamInput, 1, optValor(3).Value
    oComando.CreateParameter "@OrigenVenta", adVarChar, adParamInput, 30, Me.cboOrigenVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    
    
    
'    If rsReporte.RecordCount = 0 Then
'            MsgBox "No hay Datos para Mostrar", vbInformation, sMensaje
'            Exit Sub
'    End If
'
    Cn.Execute "delete from " & sDetalle
    
    Dim sLinea As String
    Dim i As Integer
        
    If Not (rsReporte.EOF Or rsReporte.BOF) Then
    rsReporte.MoveFirst
        Do While Not rsReporte.EOF
                If Not IsNull(rsReporte!mes) Then
                
                
                    If Calcular("select count(tcodigoproducto) as Codigo from " & sDetalle & " where tcodigoproducto ='" & IIf(IsNull(rsReporte!tCodigoProducto), "", rsReporte!tCodigoProducto) & "' and quiebre='" & IIf(IsNull(rsReporte!quiebre), "", rsReporte!quiebre) & "'", Cn) > 0 And (IIf(IsNull(rsReporte!Producto), "", rsReporte!Producto) <> "") Then
                       Isql = "update " & sDetalle & " set C" & rsReporte!mes & "= C" & rsReporte!mes & " + " & IIf(IsNull(rsReporte!Cantidad), 0, rsReporte!Cantidad) & ", M" & rsReporte!mes & "= M" & rsReporte!mes & " + " & IIf(IsNull(rsReporte!Venta), 0, rsReporte!Venta) & _
                              " where tcodigoproducto='" & IIf(IsNull(rsReporte!tCodigoProducto), "", rsReporte!tCodigoProducto) & "'"
                    Else
                       sLinea = ""
                       For i = 1 To 12
                           If rsReporte!mes <> i Then
                              sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ", "
                           End If
                       Next i
                       sLinea = sLinea & " C" & rsReporte!mes & ", M" & rsReporte!mes
                       Isql = "insert " & sDetalle & " (quiebre,tgrupo, tsubGrupo,tcodigoproducto,tproducto,tipoproducto,tagrupacion,enero,febrero,marzo,abril,mayo,junio,julio,agosto,setiembre,octubre,noviembre,diciembre," & sLinea & ") " & _
                              "values ('" & IIf(IsNull(rsReporte!quiebre), "", rsReporte!quiebre) & "','" & IIf(IsNull(rsReporte!Grupo), "", rsReporte!Grupo) & "','" & IIf(IsNull(rsReporte!SubGrupo), "", rsReporte!SubGrupo) & "','" & IIf(IsNull(rsReporte!tCodigoProducto), "", rsReporte!tCodigoProducto) & "', '" & IIf(IsNull(rsReporte!Producto), "", rsReporte!Producto) & "','" & IIf(IsNull(rsReporte!tipoProducto), "", rsReporte!tipoProducto) & "','" & IIf(IsNull(rsReporte!tagrupacion), "", rsReporte!tagrupacion) & "'" & _
                              " ," & rsReporte!Enero & "," & rsReporte!Febrero & "  , " & rsReporte!Marzo & "  , " & rsReporte!Abril & "  , " & rsReporte!Mayo & "  , " & rsReporte!Junio & "  , " & rsReporte!Julio & "  , " & rsReporte!Agosto & "  , " & rsReporte!Setiembre & "  , " & rsReporte!Octubre & "  , " & rsReporte!Noviembre & "  , " & rsReporte!Diciembre & "  ,  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & rsReporte!Cantidad & ", " & rsReporte!Venta & " )"
                       'Isql = "insert " & sDetalle & " (tgrupo, tsubGrupo,tcodigoproducto,tproducto,tipoproducto,tagrupacion " & sLinea & ") " & _
                              "values ('" & IIf(IsNull(rsReporte!Grupo), "", rsReporte!Grupo) & "','" & IIf(IsNull(rsReporte!SubGrupo), "", rsReporte!SubGrupo) & "','" & IIf(IsNull(rsReporte!tcodigoproducto), "", rsReporte!tcodigoproducto) & "', '" & IIf(IsNull(rsReporte!Producto), "", rsReporte!Producto) & "','" & IIf(IsNull(rsReporte!tipoProducto), "", rsReporte!tipoProducto) & "','" & IIf(IsNull(rsReporte!tagrupacion), "", rsReporte!tagrupacion) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & rsReporte!Cantidad & ", " & rsReporte!Venta & " )"
                    End If
                End If
                Cn.Execute Isql
            rsReporte.MoveNext
        Loop
    End If
    'Set rsReporteO = Lib.OpenRecordset("SELECT * FROM " & sDetalle & " ORDER BY TPRODUCTO", Cn)
    Set rsReporteO = Lib.OpenRecordset("select quiebre, tGrupo,tSubGrupo, tcodigoProducto, tProducto,tipoproducto,tagrupacion,C1,M1,C2,M2,C3,M3,C4, M4, C5, M5, C6, M6,C7,M7, C8,M8,C9,M9,C10,M10,C11,M11,C12,M12,enero,febrero,marzo,abril,mayo,junio,julio,agosto,setiembre,octubre,noviembre,diciembre from " & sDetalle & " order by tproducto", Cn)
    If optValor(3).Value = True Then
       Resumido.DiscardSavedData
       Resumido.Text1.SetText localConectado
       Resumido.Database.SetDataSource rsReporteO
       Resumido.Text20.SetText sRazonSocial
       Resumido.ReportTitle = "Periodo : " & Year(Me.dtpAnual.Value) & sTitulo & Chr(13) & "Expresado en Unidades"
       Resumido.Text8.SetText "Origen de datos: " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", IIf(optOrigen(8).Value = True, "Pedidos Cancelados como Cortesia", "Combinaciones"))))))
       frmEmite.CRViewer.ReportSource = Resumido
    End If
       
     If optValor(3).Value = False Then
       ResumidoMonto.DiscardSavedData
       ResumidoMonto.Text1.SetText localConectado
       ResumidoMonto.Database.SetDataSource rsReporteO
       ResumidoMonto.Text20.SetText sRazonSocial
       ResumidoMonto.ReportTitle = "Periodo : " & Year(Me.dtpAnual.Value) & sTitulo & Chr(13) & "Expresado en Nuevos Soles"
       ResumidoMonto.Text8.SetText "Origen de datos: " & IIf(optOrigen(0).Value = True, "Producción", IIf(optOrigen(1).Value = True, "Ventas", IIf(optOrigen(2).Value = True, "Cortesías", IIf(optOrigen(3).Value = True, "Cuentas Corrientes", IIf(optOrigen(6).Value = True, "Pedidos Facturados", IIf(optOrigen(8).Value = True, "Pedidos Cancelados como Cortesia", "Combinaciones"))))))
       frmEmite.CRViewer.ReportSource = ResumidoMonto
    End If
       

End Sub

Private Sub optOpcion_Click(Index As Integer)
   If Index = 7 Then
      optValor(0).Enabled = False
      optValor(1).Enabled = False
      optValor(2).Enabled = False
   Else
      optValor(0).Enabled = True
      optValor(1).Enabled = True
      optValor(2).Enabled = True
   End If
End Sub


'Public Sub Genera()
'    Screen.MousePointer = vbHourglass
'    Dim oComando As clsComando
'    Dim sFecha As String
'    Dim sAno As String
'    Dim sMes As String
'    Dim sDia As String
'    Dim j, i As Integer
'    j = 0
'
'    Set oComando = New clsComando
'    If optValor(0).value Then
'       sPrecio = "dbo.DPEDIDO.nVenta"
'       sTexto = " con Precios de Venta"
'    Else
'       sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
'       sTexto = " con Precios Netos"
'    End If
'
'    sFecha = ""
'    sAno = LTrim(str(Year(dtpAnual.value)))
'    sMes = IIf(CmbMes.ListIndex + 1 > 9, LTrim(str(CmbMes.ListIndex + 1)), "0" & LTrim(str(CmbMes.ListIndex + 1)))
'
'    Select Case CmbMes.ListIndex + 1
'           Case 2
'                If Year(dtpAnual.value) Mod 4 = 0 Then
'                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
'                Else
'                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
'                End If
'
'           Case 1, 3, 5, 7, 8, 10, 12
'                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
'
'           Case 4, 6, 9, 11
'                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
'
'    End Select
'
'     If Not oComando.CreateCmdSp("spRep_VentaFecha", Cn) Then
'      Set oComando = Nothing
'      Exit Sub
'     End If
'     oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 60, sPrecio
'     oComando.CreateParameter "@sAno", adVarChar, adParamInput, 4, sAno
'     oComando.CreateParameter "@sMes", adVarChar, adParamInput, 4, sMes
'     oComando.CreateParameter "@sFecha", adVarChar, adParamInput, 400, sFecha
'     oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
'
'   If Not oComando.GetParamOK Then
'      Set oComando = Nothing
'      Exit Sub
'   End If
'      Set RsReporte = oComando.GetSP()
'    ReporteMensual.DiscardSavedData
'    ReporteMensual.Database.SetDataSource RsReporte
'    ReporteMensual.PaperOrientation = crLandscape
'    ReporteMensual.ReportTitle = " Comparativo del  Mes de : " & CmbMes.Text & sTexto
'    ReporteMensual.Text4.SetText IIf(Trim(sBoton1) = "", "no activo", sBoton1)
'    ReporteMensual.Text5.SetText IIf(Trim(sBoton2) = "", "no activo", sBoton2)
'    ReporteMensual.Text6.SetText IIf(Trim(sBoton3) = "", "no activo", sBoton3)
'    ReporteMensual.Text7.SetText IIf(Trim(sBoton4) = "", "no activo", sBoton4)
'    ReporteMensual.Text8.SetText IIf(Trim(sBoton5) = "", "no activo", sBoton5)
'    ReporteMensual.Text20.SetText sRazonSocial
'    ReporteMensual.Text9.SetText localConectado
'    frmEmite.CRViewer.ReportSource = ReporteMensual
'End Sub


