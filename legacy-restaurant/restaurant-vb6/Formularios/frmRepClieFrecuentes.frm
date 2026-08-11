VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepClieFrecuentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Cliente Frecuente"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmRepClieFrecuentes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7905
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
      Left            =   4950
      Picture         =   "frmRepClieFrecuentes.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7410
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
      Height          =   7350
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   7815
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
         Left            =   4920
         TabIndex        =   53
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CheckBox ChkFranjaHoraria 
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
         Height          =   375
         Left            =   4920
         TabIndex        =   52
         Top             =   4920
         Width           =   1935
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
         Left            =   4950
         TabIndex        =   45
         Top             =   3570
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
         Left            =   4950
         TabIndex        =   44
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2175
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
         Left            =   4950
         TabIndex        =   43
         Top             =   4380
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CommandButton cmdBusca 
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
         Height          =   375
         Index           =   2
         Left            =   3780
         Picture         =   "frmRepClieFrecuentes.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4335
         Width           =   1005
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
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4365
         Width           =   2145
      End
      Begin VB.CheckBox ChkTipoCliente 
         Caption         =   "Todos los Tipos de Clientes"
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
         Left            =   4950
         TabIndex        =   37
         Top             =   660
         Value           =   1  'Checked
         Width           =   2700
      End
      Begin VB.CheckBox chkPedido 
         Caption         =   "Todos los Tipos de Pedidos"
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
         Left            =   4950
         TabIndex        =   34
         Top             =   1095
         Value           =   1  'Checked
         Width           =   2700
      End
      Begin VB.CheckBox chkMonto 
         Caption         =   "Sin Monto Mínimo"
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
         Left            =   4950
         TabIndex        =   33
         Top             =   2715
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   2700
         Width           =   2145
      End
      Begin VB.CheckBox chkMotorizado 
         Caption         =   "Todos los Motorizados"
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
         Left            =   4950
         TabIndex        =   27
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2265
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
         Left            =   4950
         TabIndex        =   25
         Top             =   330
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   3780
         Picture         =   "frmRepClieFrecuentes.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3105
         Width           =   1005
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
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   300
         Width           =   2145
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
         Left            =   3780
         Picture         =   "frmRepClieFrecuentes.frx":0A32
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   1005
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
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3120
         Width           =   2145
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
         Height          =   1620
         Left            =   3840
         TabIndex        =   22
         Top             =   5640
         Width           =   3915
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado x Dia"
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
            Left            =   180
            TabIndex        =   51
            Top             =   1280
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Nuevos Clientes"
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
            Left            =   180
            TabIndex        =   40
            Top             =   1020
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado por Zonas"
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
            Left            =   180
            TabIndex        =   32
            Top             =   765
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado por Cliente"
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
            Left            =   180
            TabIndex        =   10
            Top             =   255
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado por Cliente"
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
            Left            =   180
            TabIndex        =   11
            Top             =   510
            Width           =   2835
         End
      End
      Begin VB.ComboBox cboOrden 
         Height          =   315
         ItemData        =   "frmRepClieFrecuentes.frx":0B34
         Left            =   1530
         List            =   "frmRepClieFrecuentes.frx":0B41
         TabIndex        =   9
         Top             =   5610
         Width           =   2145
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
         Left            =   4950
         TabIndex        =   4
         Top             =   3150
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkZona 
         Caption         =   "Todas las Zonas"
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
         Left            =   4950
         TabIndex        =   1
         Top             =   2310
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1530
         TabIndex        =   7
         Top             =   5175
         Width           =   2145
         _ExtentX        =   3784
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
         Left            =   1530
         TabIndex        =   5
         Top             =   4755
         Width           =   2145
         _ExtentX        =   3784
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
         Height          =   345
         Left            =   3780
         TabIndex        =   6
         Top             =   4755
         Width           =   1005
         _ExtentX        =   1773
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
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   345
         Left            =   3780
         TabIndex        =   8
         Top             =   5175
         Width           =   1005
         _ExtentX        =   1773
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
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboZona 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   2295
         Width           =   2145
         _ExtentX        =   3784
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
      Begin MSDataListLib.DataCombo cboMotorizado 
         Height          =   315
         Left            =   1515
         TabIndex        =   28
         Top             =   1905
         Width           =   2145
         _ExtentX        =   3784
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
      Begin MSDataListLib.DataCombo cboPedido 
         Height          =   315
         Left            =   1530
         TabIndex        =   35
         Top             =   1100
         Width           =   2145
         _ExtentX        =   3784
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
      Begin MSDataListLib.DataCombo CboTipoCliente 
         Height          =   315
         Left            =   1545
         TabIndex        =   38
         Top             =   705
         Width           =   2145
         _ExtentX        =   3784
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
         Left            =   1530
         TabIndex        =   46
         Top             =   3555
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1530
         TabIndex        =   47
         Top             =   3960
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
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   1530
         TabIndex        =   54
         Top             =   1500
         Width           =   2145
         _ExtentX        =   3784
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
         Left            =   0
         TabIndex        =   55
         Top             =   1560
         Width           =   1515
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
         Index           =   10
         Left            =   810
         TabIndex        =   50
         Top             =   3660
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
         Index           =   9
         Left            =   480
         TabIndex        =   49
         Top             =   4050
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
         Index           =   8
         Left            =   555
         TabIndex        =   48
         Top             =   4425
         Width           =   900
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cliente :"
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
         Left            =   330
         TabIndex        =   39
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pedido :"
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
         Left            =   315
         TabIndex        =   36
         Top             =   1185
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Mínimo :"
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
         Left            =   135
         TabIndex        =   30
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorizado :"
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
         Left            =   405
         TabIndex        =   29
         Top             =   1950
         Width           =   1065
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   750
         TabIndex        =   26
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ranking por :"
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
         Left            =   300
         TabIndex        =   21
         Top             =   5670
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   345
         TabIndex        =   20
         Top             =   5250
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   19
         Top             =   4830
         Width           =   1230
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         TabIndex        =   18
         Top             =   3195
         Width           =   630
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona :"
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
         Left            =   900
         TabIndex        =   17
         Top             =   2355
         Width           =   570
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
      Left            =   6405
      Picture         =   "frmRepClieFrecuentes.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7410
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
      Left            =   3495
      Picture         =   "frmRepClieFrecuentes.frx":0C5E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7410
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
      Left            =   2040
      Picture         =   "frmRepClieFrecuentes.frx":1190
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7410
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   225
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepClieFrecuentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim Rsturno As Recordset
Dim RsZona As Recordset
Dim RsMotorizado As Recordset
Dim RsPedido As Recordset
Dim RsTipoCliente As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsProducto As Recordset
Dim RsOrigen As Recordset

Dim Detallado As New dsrDeliveryDetallado
Dim DetalladoxDia As New dsrDeliveryDetalladoxDia
Dim Grupo As New dsrDeliveryGrupo
Dim sCriterio As String
Dim xCriterio As String
Dim scriterio1 As String
Dim sCriterio2 As String
Dim xMonto1 As String
Dim xMonto2 As String
Dim sTitulo As String
Dim sProducto   As String

Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboZona
         Isql = "Select * from vZona"
         Set RsZona = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsZona
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboMotorizado
         Isql = "Select * from vMotorizado"
         Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotorizado
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboPedido
         Isql = "select * from vTipoPedido"
         Set RsPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsPedido
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
    
    With cboTipoCliente
         Isql = "select tdetallado as descripcion,tcodigo as codigo from ttabla where ttabla='TIPOCLIENTEFRECUENTE'" ' and tcodigo <>'00'"
         Set RsTipoCliente = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoCliente
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



Private Sub cboPedido_Change()
    Me.cboOrigenVenta.BoundText = ""
    With cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta  " & IIf(cboPedido.BoundText = "", "", "where CodCanalVenta = '" & cboPedido.BoundText & "' and Activo=1") & " order by Descripcion "
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
    End With
End Sub
Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkproducto_Click()
   If chkProducto.Value = 1 Then
      sProducto = ""
      txtProducto.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
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
Private Sub chkSubGrupo_Click()
   If chkSubGrupo.Value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub

Private Sub ChkTipoCliente_Click()
   If ChkTipoCliente.Value = 1 Then
      cboTipoCliente.Enabled = False
      cboTipoCliente.Text = ""
   Else
      cboTipoCliente.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
Select Case Index
   Case Is = 0
      Isql = "select Codigo, isnull(tApellido,'') + ' ' + isnull(tNombre,'') as Descripcion, tTelefono  from vdelivery order by Descripcion"
      frmBusca.cboCriterio.Enabled = True
      frmBusca.nPredeterm = 1
      Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                             "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                             "Teléfono", 2, "tTelefono", 1500, 0, 0, "")
      frmBusca.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      sCliente = sCodigo
      txtCliente.Text = sDescrip

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
          
   End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""
   scriterio1 = ""
   sCriterio2 = ""
   xMonto1 = ""
   xMonto2 = ""
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
         xCriterio = sCriterio
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
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
      xCriterio = sCriterio
   End If
   
   If chkPedido.Value = 0 Then
      If cboPedido.Text = "" Then
         MsgBox "Debe escoger un tipo de pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboPedido.BoundText & "'"
      xCriterio = sCriterio
   End If
    If Me.chkOrigenVenta.Value = 0 Then
      If Me.cboOrigenVenta.Text = "" Then
         MsgBox "Debe escoger un Origen Venta", vbCritical, sMensaje
         Exit Sub
      End If
        sCriterio = sCriterio & " and MPEDIDO.CodigoOrigenVenta ='" & cboOrigenVenta.BoundText & "'"
        xCriterio = sCriterio
   End If
   If ChkTipoCliente.Value = 0 Then
      If cboTipoCliente.Text = "" Then
         MsgBox "Debe escoger un tipo de Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      scriterio1 = " and vDelivery.tTipoCliente ='" & cboTipoCliente.BoundText & "'"
  End If
   
  If chkMotorizado.Value = 0 Then
     If cboMotorizado.Text = "" Then
        MsgBox "Debe escoger un Motorizado", vbCritical, sMensaje
        Exit Sub
     End If
     sCriterio = sCriterio & " and MPEDIDO.tMotorizado ='" & cboMotorizado.BoundText & "'"
     xCriterio = sCriterio
  End If
   
  If chkCliente.Value = 0 Then
     If txtCliente.Text = "" Then
        MsgBox "Debe escoger un cliente", vbCritical, sMensaje
        Exit Sub
     End If
     sCriterio = sCriterio & " and MPEDIDO.tClienteDelivery ='" & sCliente & "'"
     xCriterio = sCriterio
  End If
   
  If chkMonto.Value = 0 Then
     If CDbl(txtMonto) < 0 Then
        MsgBox "El monto debe ser mayor a cero", vbCritical, sMensaje
        Exit Sub
     End If
     xMonto1 = " and Monto >= " & CDbl(txtMonto.Text)
     xMonto2 = " HAVING SUM(dbo.DPEDIDO.nVenta) >= " & CDbl(txtMonto.Text)
  End If
   
  If chkZona.Value = 0 Then
     If cboZona.Text = "" Then
        MsgBox "Debe escoger una Zona", vbCritical, sMensaje
        Exit Sub
     End If
     sCriterio = sCriterio & " and vDelivery.Zona ='" & cboZona.Text & "'"
  End If
   
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio2 = sCriterio2 & " and vProducto.Grupo='" & cboGrupo.Text & "'"
   End If
   
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio2 = sCriterio2 & " and subGrupo='" & cboSubGrupo.Text & "'"
   End If
   
   If chkProducto.Value = 0 Then
      If sProducto = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio2 = sCriterio2 & " and DPEDIDO.tCodigoProducto ='" & sProducto & "'"
   End If
           
           
  Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = True
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
               Else
                   If optOpcion(3).Value = True Then
                        frmEmite.CRViewer.DisplayGroupTree = True
                        DetalladoxDia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        DetalladoxDia.PaperOrientation = crPortrait
                   Else
                        frmEmite.CRViewer.DisplayGroupTree = False
                        Grupo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Grupo.PaperOrientation = crPortrait
                   End If
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
               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
               Else
                  If optOpcion(3).Value = True Then
                     DetalladoxDia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                     DetalladoxDia.PaperOrientation = crPortrait
                     DetalladoxDia.PrintOut
                  Else
                     Grupo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                     Grupo.PaperOrientation = crPortrait
                     Grupo.PrintOut
                  End If
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
               If optOpcion(0).Value = True Then
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
                  If optOpcion(3).Value = True Then
                     DetalladoxDia.ExportOptions.FormatType = 21
                     DetalladoxDia.ExportOptions.DestinationType = 1
                     cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                     cmdSave.ShowSave
                     If cmdSave.FileName = "" Then
                        Exit Sub
                     End If
                     DetalladoxDia.ExportOptions.DiskFileName = cmdSave.FileName
                     DetalladoxDia.Export False
                  Else
                     Grupo.ExportOptions.FormatType = 21
                     Grupo.ExportOptions.DestinationType = 1
                     cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                     cmdSave.ShowSave
                     If cmdSave.FileName = "" Then
                        Exit Sub
                     End If
                     Grupo.ExportOptions.DiskFileName = cmdSave.FileName
                     Grupo.Export False
                  End If
                                
               End If
   End Select

End Sub

Private Sub chkCliente_Click()
   If chkCliente.Value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkMonto_Click()
   If chkMonto.Value = 1 Then
      txtMonto.Enabled = False
      txtMonto.Text = "0.00"
   Else
      txtMonto.Enabled = True
   End If
End Sub

Private Sub chkPedido_Click()
   If chkPedido.Value = 1 Then
      cboPedido.Enabled = False
      cboPedido.Text = ""
   Else
      cboPedido.Enabled = True
   End If
End Sub

Private Sub chkMotorizado_Click()
   If chkMotorizado.Value = 1 Then
      cboMotorizado.Enabled = False
      cboMotorizado.Text = ""
   Else
      cboMotorizado.Enabled = True
   End If
End Sub

Private Sub chkZona_Click()
   If chkZona.Value = 1 Then
      cboZona.Enabled = False
      cboZona.Text = ""
   Else
      cboZona.Enabled = True
   End If
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
      sTurno = ""
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
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    sTurno = ""
    cboZona.Enabled = False
    cboPedido.Enabled = False
    cboMotorizado.Enabled = False
    cboZona.Text = ""
    cboPedido.Text = ""
    cboMotorizado.Text = ""
    txtMonto.Text = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsZona = Nothing
   Set RsPedido = Nothing
   Set RsMotorizado = Nothing
End Sub

Public Sub Genera()
   Dim oComando As clsComando
    Dim tipopedido As String
   Dim sOrden As String
   Dim xOrden As String
   Screen.MousePointer = vbHourglass
   
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_RepClieFrecuentes", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
   Select Case cboOrden.Text
          Case Is = "Descripción"
               sOrden = "Cliente"
               xOrden = "Grupo"
          Case Is = "Monto"
               sOrden = "T1.Monto DESC, Cliente"
               xOrden = "Monto DESC, Grupo"
          Case Else
               sOrden = "T1.Pedido DESC, T1.Monto DESC, Cliente"
               xOrden = "Pedido DESC, Grupo"
   End Select
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
       If Me.cboPedido.Text <> "" And Me.chkPedido.Value = 0 Then
        tipopedido = cboPedido.BoundText
        Else
        tipopedido = ""
    End If
    oComando.CreateParameter "@flagFranjaHoraria", adBoolean, adParamInput, 1, IIf(ChkFranjaHoraria.Value = 1, True, False)
    oComando.CreateParameter "@flagDetaxClie", adBoolean, adParamInput, 1, optOpcion(0).Value
    oComando.CreateParameter "@flagConsoxClie", adBoolean, adParamInput, 1, optOpcion(1).Value
    oComando.CreateParameter "@flagConsoxZona", adBoolean, adParamInput, 1, optOpcion(2).Value
    oComando.CreateParameter "@flagNuevxClie", adBoolean, adParamInput, 1, optOpcion(4).Value
    oComando.CreateParameter "@flagDetaxDia", adBoolean, adParamInput, 1, optOpcion(3).Value
    oComando.CreateParameter "@flagMonto", adBoolean, adParamInput, 1, Me.chkMonto.Value
    oComando.CreateParameter "@flagTurnoOFecha", adBoolean, adParamInput, 1, Me.chkTurno.Value
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tTipoCliente", adVarChar, adParamInput, 30, cboTipoCliente.BoundText
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 30, cboMotorizado.BoundText
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 30, sCliente
    oComando.CreateParameter "@tZona", adVarChar, adParamInput, 30, cboZona.Text
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tcodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@tMonto", adDouble, adParamInput, 10, CDbl(txtMonto.Text)
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 50, sOrden
    oComando.CreateParameter "@xOrden", adVarChar, adParamInput, 50, xOrden
    oComando.CreateParameter "@turno", adVarChar, adParamInput, 20, sTurno
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@OrigenVenta", adVarChar, adParamInput, 30, Me.cboOrigenVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
   
   
   If optOpcion(0).Value = True Then
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource rsReporte
      Detallado.Text13.SetText sRazonSocial
      Detallado.Text14.SetText localConectado
      Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
      frmEmite.CRViewer.ReportSource = Detallado
   Else
   If optOpcion(3).Value = True Then
      DetalladoxDia.DiscardSavedData
      DetalladoxDia.Database.SetDataSource rsReporte
      DetalladoxDia.Text13.SetText sRazonSocial
      DetalladoxDia.Text14.SetText localConectado
      DetalladoxDia.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
      frmEmite.CRViewer.ReportSource = DetalladoxDia
       Else
                If optOpcion(0).Value = False Then
                    Grupo.DiscardSavedData
                    Grupo.Database.SetDataSource rsReporte
                    Grupo.Text13.SetText sRazonSocial
                    Grupo.Text14.SetText localConectado
                    If optOpcion(1).Value = True Then
                        Grupo.Text6.SetText "Cliente"
                        Grupo.Text10.SetText "Reporte Agrupado por Clientes"
                    End If
                    If optOpcion(2) Then
                        Grupo.Text6.SetText "Zona"
                        Grupo.Text5.SetText ""
                        Grupo.Text4.SetText ""
                        Grupo.Text10.SetText "Reporte Agrupado por Zonas"
                    End If
                    If optOpcion(4) Then
                        Grupo.Text6.SetText "Cliente"
                        Grupo.Text10.SetText "Reporte de Nuevos Clientes"
                    End If
                    Grupo.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
                    frmEmite.CRViewer.ReportSource = Grupo
                 End If
        End If
   End If
End Sub



Private Sub txtMonto_gotfocus()
   txtMonto.SelLength = Len(Trim(txtMonto.Text))
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtMonto_LostFocus()
   txtMonto.Text = Format(txtMonto.Text, "##,###,##0.00")
End Sub
