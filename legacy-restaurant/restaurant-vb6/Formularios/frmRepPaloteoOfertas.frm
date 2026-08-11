VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepPaloteoOfertas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Ofertas"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepPaloteoOfertas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7575
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
      Picture         =   "frmRepPaloteoOfertas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6420
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
      Height          =   6375
      Left            =   30
      TabIndex        =   31
      Top             =   0
      Width           =   7500
      Begin VB.CheckBox chkOferta 
         Caption         =   "Todas las Ofertas"
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
         TabIndex        =   0
         Top             =   315
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo de Oferta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   2520
         TabIndex        =   43
         Top             =   4680
         Width           =   1680
         Begin VB.OptionButton optValor 
            Caption         =   "Manual"
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
            TabIndex        =   21
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Automática"
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
            TabIndex        =   22
            Top             =   585
            Width           =   1455
         End
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
         TabIndex        =   10
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2265
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
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   8
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
         TabIndex        =   6
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
         Height          =   345
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3120
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
         Picture         =   "frmRepPaloteoOfertas.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3120
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
         Height          =   1635
         Left            =   4440
         TabIndex        =   35
         Top             =   4560
         Width           =   2925
         Begin VB.OptionButton optOpcion 
            Caption         =   "Producto Por Oferta Detallado"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   720
            Width           =   2685
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Producto Por Oferta Resumido"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   2685
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Ofertas Por Producto"
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   25
            Top             =   1080
            Width           =   2640
         End
      End
      Begin VB.CheckBox chkMozo 
         Caption         =   "Todos los Usuarios"
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
         TabIndex        =   4
         Top             =   1035
         Value           =   1  'Checked
         Width           =   2265
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
         TabIndex        =   2
         Top             =   675
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
         Height          =   315
         Left            =   5130
         TabIndex        =   16
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1890
         TabIndex        =   19
         Top             =   4005
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
         Left            =   1890
         TabIndex        =   17
         Top             =   3615
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
         Left            =   5160
         TabIndex        =   18
         Top             =   3600
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   3990
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   600
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
         TabIndex        =   3
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
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
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   1665
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
         TabIndex        =   11
         Top             =   2385
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
         TabIndex        =   13
         Top             =   2745
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
         TabIndex        =   9
         Top             =   2025
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
      Begin MSDataListLib.DataCombo cboOferta 
         Height          =   315
         Left            =   1860
         TabIndex        =   45
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oferta :"
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
         Left            =   1080
         TabIndex        =   44
         Top             =   360
         Width           =   660
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
         Left            =   600
         TabIndex        =   42
         Top             =   3690
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
         Left            =   705
         TabIndex        =   41
         Top             =   4080
         Width           =   1125
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
         TabIndex        =   40
         Top             =   2145
         Width           =   1500
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   1440
         Width           =   1545
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
         Index           =   9
         Left            =   960
         TabIndex        =   34
         Top             =   1080
         Width           =   780
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
         TabIndex        =   33
         Top             =   720
         Width           =   1425
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
         Left            =   1200
         TabIndex        =   32
         Top             =   3240
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
      Picture         =   "frmRepPaloteoOfertas.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6420
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
      Picture         =   "frmRepPaloteoOfertas.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6420
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
      Picture         =   "frmRepPaloteoOfertas.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6420
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   810
      Top             =   6390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepPaloteoOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim rsoferta As Recordset
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

Dim sCriterio   As String
Dim xCriterio   As String
Dim sProducto   As String
Dim sTurno      As String
Dim sPropiedad  As String
Dim sTitulo     As String
Dim sTexto      As String


Dim Detallado As New dsrPaloteoOfertaD
Dim Resumido As New dsrPaloteoOfertaR
Dim Producto As New dsrPaloteoOfertaP
Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboOferta
        Isql = "select distinct * from voferta where lactivo=1"
        Set rsoferta = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = rsoferta
             .DataField = "tnombre"
             .ListField = "tnombre"
             .BoundColumn = "toferta"
    End With
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboMozo
         Isql = "Select * from TUSUARIO where lActivo = 1 and tGrupoUsuario<>'00' order by tCodigoUsuario"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "tCodigoUsuario"
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
   If ChkCaja.Value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

Private Sub chkOferta_Click()
   If chkOferta.Value = 1 Then
      cboOferta.Enabled = False
      cboOferta.Text = ""
   Else
      cboOferta.Enabled = True
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
   
   If chkArea.Value = 0 Then
      If cboArea.Text = "" Then
         MsgBox "Debe escoger un Area", vbCritical, sMensaje
        Exit Sub
      End If
      sCriterio = sCriterio & " and Area ='" & cboArea.Text & "'"
   End If
      
   If ChkCaja.Value = 0 Then
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
               ElseIf optOpcion(1).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
               ElseIf optOpcion(2).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Producto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Producto.PaperOrientation = crPortrait
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
               ElseIf optOpcion(1).Value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               ElseIf optOpcion(2).Value = True Then
                  Producto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Producto.PaperOrientation = crPortrait
                  Producto.PrintOut
               End If
'
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
               ElseIf optOpcion(1).Value = True Then
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
               ElseIf optOpcion(2).Value = True Then
                  Producto.ExportOptions.FormatType = 21
                  Producto.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Producto.ExportOptions.DiskFileName = cmdSave.FileName
                  Producto.Export False
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
              ' txtProducto.Text = sDescrip
          
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
''                                                      "Identificador", 2, "Codigo", 1500, 0, 0, "", _
''                                                      "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
''               frmBusca.Show vbModal
''               If Not wEnter Then
''                  Exit Sub
''               End If
''        '       sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
'          '     Isql = "select codigo, descripcion,nprecio from vpropiedad where toperador like '" & cboOperador.BoundText & "%' order by 2"
'               frmBusca.nPredeterm = 1
'               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "codigo", 1200, 0, 0, "", _
'                                                      "Propiedad", 2, "descripcion", 4000, 0, 0, "", _
'                                                      "Precio", 2, "nprecio", 1200, 0, 0, "")
'               frmBusca.Show vbModal
'               If Not wEnter Then
'                  Exit Sub
'               End If
'               sPropiedad = sCodigo
'               Me.txtPropiedad.Text = sDescrip
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
    cboUnidadNegocio.Enabled = False
    cboUnidadNegocio.Text = ""
    cboOferta.Enabled = False
    cboOferta.Text = ""
  '  cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
   ' cmdBusca(2).Enabled = False
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
    Dim tipopedido As String
    Dim Tipo As String
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Paloteooferta", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    If optOpcion(1).Value = True Then: Tipo = "2"
    If optOpcion(0).Value = True Then: Tipo = "1"
    If optOpcion(2).Value = True Then: Tipo = "3"
        If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.Value = 0 Then
        tipopedido = cboTipoPedido.BoundText
        Else
        tipopedido = ""
    End If
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    oComando.CreateParameter "@flagTipoReporte", adVarChar, adParamInput, 1, Tipo
    oComando.CreateParameter "@flagTurnoFecha", adBoolean, adParamInput, 1, Me.chkTurno.Value
    oComando.CreateParameter "@flagAutoManual", adBoolean, adParamInput, 1, Me.optValor(1).Value
    oComando.CreateParameter "@soFERTA", adVarChar, adParamInput, 10, Me.cboOferta.BoundText
    oComando.CreateParameter "@sTipoPedido", adVarChar, adParamInput, 10, tipopedido
    oComando.CreateParameter "@sMozo", adVarChar, adParamInput, 10, cboMozo.Text
    oComando.CreateParameter "@sArea", adVarChar, adParamInput, 10, cboArea.BoundText
    oComando.CreateParameter "@sCaja", adVarChar, adParamInput, 10, cboCaja.BoundText
    oComando.CreateParameter "@sUNegocio", adVarChar, adParamInput, 10, Me.cboUnidadNegocio.BoundText
    oComando.CreateParameter "@sGrupo", adVarChar, adParamInput, 10, Me.cboGrupo.BoundText
    oComando.CreateParameter "@sSubGrupo", adVarChar, adParamInput, 10, Me.cboSubGrupo.BoundText
    oComando.CreateParameter "@sTurno", adVarChar, adParamInput, 20, sTurno
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    If Tipo = "1" Then ' detallado
        Detallado.DiscardSavedData
        Detallado.Database.SetDataSource rsReporte
        Detallado.Text25.SetText sTitulo
        Detallado.Text30.SetText localConectado
        Detallado.Text15.SetText sRazonSocial
        Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
        frmEmite.CRViewer.ReportSource = Detallado
    ElseIf Tipo = "2" Then ' resumido
        Resumido.DiscardSavedData
        Resumido.Database.SetDataSource rsReporte
        Resumido.Text25.SetText sTitulo
        Resumido.Text30.SetText localConectado
        Resumido.Text15.SetText sRazonSocial
        Resumido.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
        frmEmite.CRViewer.ReportSource = Resumido
    ElseIf Tipo = "3" Then ' producto x oferta
        Producto.DiscardSavedData
        Producto.Database.SetDataSource rsReporte
        Producto.Text25.SetText sTitulo
        Producto.Text30.SetText localConectado
        Producto.Text15.SetText sRazonSocial
        Producto.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
        frmEmite.CRViewer.ReportSource = Producto
    End If
End Sub

