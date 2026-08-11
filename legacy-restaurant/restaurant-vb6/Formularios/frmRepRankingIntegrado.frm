VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepRankingIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ranking de Producción"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   Icon            =   "frmRepRankingIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   12450
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
      Left            =   5498
      Picture         =   "frmRepRankingIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
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
      Left            =   8408
      Picture         =   "frmRepRankingIntegrado.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
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
      Left            =   4043
      Picture         =   "frmRepRankingIntegrado.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
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
      Left            =   2588
      Picture         =   "frmRepRankingIntegrado.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
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
      Left            =   6953
      Picture         =   "frmRepRankingIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CMDSAVE 
      Left            =   600
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   5550
      Left            =   30
      TabIndex        =   33
      Top             =   30
      Width           =   12375
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
         Height          =   5475
         Left            =   7800
         TabIndex        =   44
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar Todo"
            Height          =   360
            Left            =   2880
            TabIndex        =   32
            Top             =   5040
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4815
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   8493
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
         Left            =   5040
         TabIndex        =   43
         Top             =   3150
         Width           =   2325
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
            TabIndex        =   28
            Top             =   1388
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   30
            Top             =   1935
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   29
            Top             =   1660
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
            TabIndex        =   27
            Top             =   1116
            Width           =   2025
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
            TabIndex        =   26
            Top             =   844
            Width           =   1935
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
            TabIndex        =   25
            Top             =   572
            Width           =   1935
         End
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
            TabIndex        =   24
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
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
         Height          =   1605
         Left            =   90
         TabIndex        =   42
         Top             =   3150
         Width           =   1680
         Begin VB.OptionButton optValor 
            Caption         =   "Comparativo"
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
            Left            =   90
            TabIndex        =   21
            Top             =   1215
            Width           =   1455
         End
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
            Top             =   900
            Width           =   1455
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
         Height          =   1605
         Left            =   1860
         TabIndex        =   41
         Top             =   3150
         Width           =   3135
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
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
            TabIndex        =   22
            Top             =   270
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
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
            TabIndex        =   23
            Top             =   585
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.ComboBox cboOrden 
         Height          =   315
         ItemData        =   "frmRepRankingIntegrado.frx":1D0A
         Left            =   1800
         List            =   "frmRepRankingIntegrado.frx":1D14
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2700
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
         Left            =   5130
         TabIndex        =   12
         Top             =   1455
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
         TabIndex        =   8
         Top             =   645
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
         TabIndex        =   10
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1875
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
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Top             =   2220
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
         Format          =   21037057
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1800
         TabIndex        =   13
         Top             =   1800
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
         Format          =   21037057
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   5070
         TabIndex        =   14
         Top             =   1815
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
         Format          =   21037059
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5070
         TabIndex        =   16
         Top             =   2220
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
         Format          =   21037059
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   240
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   630
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
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1035
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
         Left            =   1800
         TabIndex        =   11
         Top             =   1425
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
      Begin VB.Label Label 
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
         Left            =   615
         TabIndex        =   40
         Top             =   2730
         Width           =   1170
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
         TabIndex        =   39
         Top             =   1500
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
         TabIndex        =   38
         Top             =   690
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
         TabIndex        =   37
         Top             =   1095
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
         TabIndex        =   36
         Top             =   2355
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
         TabIndex        =   35
         Top             =   1950
         Width           =   1230
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
         TabIndex        =   34
         Top             =   285
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmRepRankingIntegrado"
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
Dim RsTipoPedido As Recordset
Dim RsTipo As Recordset
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim rankingL As New dsrRankingIntegradoL
Dim rankingC As New dsrRankingIntegradoComparativo
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
   Dim r, p
   Cn.Execute "delete from " & sNombre
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
   ' MsgBox dtpFecFin.value - dtpFecIni
   

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
               If RsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               If optValor(3).value = False Then
                    If optOpcion(0).value = True Then
                       frmEmite.CRViewer.DisplayGroupTree = True
                       rankingL.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingL.PaperOrientation = crPortrait
                    Else
                       frmEmite.CRViewer.DisplayGroupTree = True
                       rankingL.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingL.PaperOrientation = crPortrait
                    End If
                Else
                       frmEmite.CRViewer.DisplayGroupTree = True
                       rankingC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingC.PaperOrientation = crPortrait
                End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               If optValor(3).value = False Then
                    If optOpcion(0).value = True Then
                       rankingL.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingL.PaperOrientation = crPortrait
                       rankingL.PrintOut
                    Else
                       rankingL.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingL.PaperOrientation = crPortrait
                       rankingL.PrintOut
                    End If
                Else
                       rankingC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                       rankingC.PaperOrientation = crPortrait
                       rankingC.PrintOut
                End If
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optValor(3).value = False Then
                        If optOpcion(0).value = True Then
                           rankingL.ExportOptions.FormatType = 21
                           rankingL.ExportOptions.DestinationType = 1
                           CMDSAVE.Filter = "Libro de Microsoft Excel|*.xls"
                           CMDSAVE.ShowSave
                           If CMDSAVE.FileName = "" Then
                              Exit Sub
                           End If
                           rankingL.ExportOptions.DiskFileName = CMDSAVE.FileName
                           rankingL.Export False
                        Else
                           rankingL.ExportOptions.FormatType = 21
                           rankingL.ExportOptions.DestinationType = 1
                           CMDSAVE.Filter = "Libro de Microsoft Excel|*.xls"
                           CMDSAVE.ShowSave
                           If CMDSAVE.FileName = "" Then
                              Exit Sub
                           End If
                           rankingL.ExportOptions.DiskFileName = CMDSAVE.FileName
                           rankingL.Export False
                        End If
                Else
                           rankingC.ExportOptions.FormatType = 21
                           rankingC.ExportOptions.DestinationType = 1
                           CMDSAVE.Filter = "Libro de Microsoft Excel|*.xls"
                           CMDSAVE.ShowSave
                           If CMDSAVE.FileName = "" Then
                              Exit Sub
                           End If
                           rankingC.ExportOptions.DiskFileName = CMDSAVE.FileName
                           rankingC.Export False
                End If
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
    sNombre = dbTemporal(sCaja, 11, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "Grupo", "nVarChar(150)", "SubGrupo", "nVarChar(150)", "TipoProducto", "nVarChar(150)", "Producto", "nVarChar(150)", "Cantidad", "Float", "Venta", "Float", "Neto", "Float", "Costo", "Float")
        
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
   ' configuraListaServidores
   obtieneListaServidores
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
   Set RsTipo = Nothing
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
    On Error GoTo sigue
   Dim oComando As New clsComando
   Dim conectado As Boolean
   Dim sCombo As String
   Dim sOrden As String
   Dim RsDatos As New ADODB.Recordset
   Dim RsCabecera As New ADODB.Recordset
   Dim w, z As Long
   Dim sPrecio As String
   Dim sNeto As String
   Dim sCosto As String
   Dim sCostoPropiedad As String
   Dim sCostoComboPropiedad As String
   Dim cs As String
   Dim cl As String
   Dim Cabecera As String
   Dim sector As String
   Dim llocal As String
   Set oComando = New clsComando
   Screen.MousePointer = vbHourglass

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
                          sTexto = "Precio Costo"
                       ElseIf optValor(3).value Then
                          sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                       End If
                          sNeto = " dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad "
                          sCosto = " (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) *  dbo.DPEDIDO.nCantidad "
                       
                    End If
                
                   'Venta
                   If optOrigen(1).value = True Then
                      If optValor(0).value Then
                         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                         sTexto = "Precio Venta"
                      ElseIf optValor(1).value Then
                         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                         sTexto = "Precio Neto"
                      ElseIf optValor(2).value Then
                         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                         sTexto = "Precio Costo"
                      Else
                         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                      End If
                         sNeto = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad "
                         sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                      
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
                         sTexto = "Precio Costo"
                      Else
                        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                      End If
                        sNeto = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad "
                        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                        
                      
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
                        sTexto = "Precio Costo"
                     Else
                        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                     End If
                        sNeto = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad "
                        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                     
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
                         sTexto = "Precio Costo"
                      Else
                        sPrecio = "dbo.CPEDIDO.nVenta * dbo.DPEDIDO.nCantidad"
                      End If
                        sNeto = "dbo.CPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad "
                        sCosto = "(isnull(dbo.vProducto.nInsumo,0)+isnull(dbo.vProducto.nGasto,0)+isnull(dbo.vProducto.nManoObra,0)) * dbo.CPEDIDO.nCantidad "
                      
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
                        sTexto = "Precio Costo"
                     Else
                        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                     End If
                        sNeto = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad "
                        sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                     
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
                         sTexto = "Precio Costo"
                      Else
                         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
                      End If
                         sNeto = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad "
                         sCosto = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad "
                      
                   End If
                   sOrden = ""
                    Select Case cboOrden.Text
                           Case Is = "Monto"
                              sOrden = " SUM(Venta) DESC "
                         Case Else
                              sOrden = " SUM(Cantidad) DESC, sum(venta) desc "
                    End Select
                    
     For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
               
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
                
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spRep_RankingIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                        Set oComando = Nothing
                        Exit Sub
                End If
                  
                        fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
                        fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
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
                        oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 200, sOrden
                        oComando.CreateParameter "@sBoton2", adVarChar, adParamInput, 30, sBoton2
                        oComando.CreateParameter "@sBoton3", adVarChar, adParamInput, 30, sBoton3
                        oComando.CreateParameter "@sBoton4", adVarChar, adParamInput, 30, sBoton4
                        oComando.CreateParameter "@sBoton5", adVarChar, adParamInput, 30, sBoton5
                        oComando.CreateParameter "@sneto", adVarChar, adParamInput, 300, sNeto
                        oComando.CreateParameter "@scosto", adVarChar, adParamInput, 300, sCosto
                        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                        If Not oComando.GetParamOK Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                        
                        cl = lstServidores.ListItems.Item(w)
                        sector = lstServidores.ListItems.Item(w).SubItems(1)
                        llocal = lstServidores.ListItems.Item(w).SubItems(2)
                        Set RsDatos = oComando.GetSP()
                        If Not (RsDatos.EOF Or RsDatos.BOF) Then
                          RsDatos.MoveFirst
                          For z = 0 To RsDatos.RecordCount - 1
                                          Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(2) & "','" & RsDatos.Fields(3) & "'," & RsDatos.Fields(4) & "," & RsDatos.Fields(5) & "," & RsDatos.Fields(6) & "," & RsDatos.Fields(7) & ")"
                                         'Debug.Print Isql
                                          Cn.Execute Isql
                            RsDatos.MoveNext
                        Next z
                        
                 End If
        End If
sigue:
    Next w
    If optValor(3).value = False Then
        If optOpcion(0).value = True Then ' sector
                            Isql = " SELECT GRUPO , SUBGRUPO,tipoproducto , producto,sum(cantidad) as cantidad, sum(venta) as valor  from " & sNombre & " group by  GRUPO , SUBGRUPO,tipoproducto , producto order by   " & sOrden
                                Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                                Isql = "select sector from " & sNombre & " group by sector"
                                Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
                                If Not (RsCabecera.EOF Or RsCabecera.BOF) Then
                                    RsCabecera.MoveFirst
                                     For w = 0 To RsCabecera.RecordCount - 1
                                        Cabecera = Cabecera & RsCabecera.Fields(0) & " / "
                                        RsCabecera.MoveNext
                                        Next w
                                End If
                                If Cabecera = "" Then: Cabecera = "///"
                            rankingL.DiscardSavedData
                            rankingL.Database.SetDataSource RsReporte
                     '       rankingL.Text3.SetText "Ranking de Producción"
                            rankingL.Text20.SetText sRazonSocial
                            rankingL.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos")
                            rankingL.Text9.SetText "Datos de Sector: " & Mid(Cabecera, 1, Len(Cabecera) - 2)
                            rankingL.Text6.SetText "Ranking Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Considerando el : " & IIf(optValor(0).value = True, optValor(0).Caption, IIf(optValor(1).value = True, optValor(1).Caption, optValor(2).Caption))
                            rankingL.Text7.SetText "Intervalo Analizado del : " & Format(Me.dtpFecIni.value, "dd/MM/yyyy") & " al " & Format(Me.dtpFecFin.value, "dd/MM/yyyy")
                            frmEmite.CRViewer.ReportSource = rankingL
        Else
                         Isql = " SELECT GRUPO , SUBGRUPO,tipoproducto , producto,sum(cantidad) as cantidad, sum(venta) as valor  from " & sNombre & " group by  GRUPO , SUBGRUPO,tipoproducto , producto order by   " & sOrden
                         Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                          For w = 1 To lstServidores.ListItems.Count
                                If lstServidores.ListItems.Item(w).Checked = True Then
                                        Cabecera = Cabecera & lstServidores.ListItems.Item(w).SubItems(2) & " / "
                                End If
                          Next w
                                
                          If Cabecera = "" Then Cabecera = "///"
                          rankingL.DiscardSavedData
                          rankingL.Database.SetDataSource RsReporte
                          rankingL.Text20.SetText sRazonSocial
                          rankingL.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos")
                          rankingL.Text9.SetText "Datos de Local: " & IIf(Cabecera <> "", Mid(Cabecera, 1, Len(Cabecera) - 2), "d")
                          rankingL.Text6.SetText "Ranking Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Considerando el : " & IIf(optValor(0).value = True, optValor(0).Caption, IIf(optValor(1).value = True, optValor(1).Caption, optValor(2).Caption))
                          rankingL.Text7.SetText "Intervalo Analizado del : " & Format(Me.dtpFecIni.value, "dd/MM/yyyy") & " al " & Format(Me.dtpFecFin.value, "dd/MM/yyyy")
                          frmEmite.CRViewer.ReportSource = rankingL
        End If
    Else
        If optOpcion(0).value = True Then ' sector
                            Isql = " SELECT GRUPO , SUBGRUPO,tipoproducto , producto,sum(cantidad) as cantidad, sum(venta) as venta,sum(neto) as neto,sum(costo) as costo  from " & sNombre & " group by  GRUPO , SUBGRUPO,tipoproducto , producto order by  " & sOrden
                                Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                                Isql = "select sector from " & sNombre & " group by sector"
                                Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
                                If Not (RsCabecera.EOF Or RsCabecera.BOF) Then
                                    RsCabecera.MoveFirst
                                     For w = 0 To RsCabecera.RecordCount - 1
                                        Cabecera = Cabecera & RsCabecera.Fields(0) & " / "
                                        RsCabecera.MoveNext
                                        Next w
                                End If
                                If Cabecera = "" Then: Cabecera = "///"
                            rankingC.DiscardSavedData
                            rankingC.Database.SetDataSource RsReporte
                     '       rankingL.Text3.SetText "Ranking de Producción"
                            rankingC.Text20.SetText sRazonSocial
                            rankingC.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos")
                            rankingC.Text9.SetText "Datos de Sector: " & Mid(Cabecera, 1, Len(Cabecera) - 2)
                            rankingC.Text6.SetText "Ranking Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Considerando el : " & IIf(optValor(0).value = True, optValor(0).Caption, IIf(optValor(1).value = True, optValor(1).Caption, optValor(2).Caption))
                            rankingC.Text7.SetText "Intervalo Analizado del : " & Format(Me.dtpFecIni.value, "dd/MM/yyyy") & " al " & Format(Me.dtpFecFin.value, "dd/MM/yyyy")
                            frmEmite.CRViewer.ReportSource = rankingC
        Else
                         Isql = " SELECT GRUPO , SUBGRUPO,tipoproducto , producto,sum(cantidad) as cantidad, sum(venta) as venta,sum(neto) as neto,sum(costo) as costo from " & sNombre & " group by  GRUPO , SUBGRUPO,tipoproducto , producto order by  " & sOrden
                         Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                          For w = 1 To lstServidores.ListItems.Count
                                If lstServidores.ListItems.Item(w).Checked = True Then
                                        Cabecera = Cabecera & lstServidores.ListItems.Item(w).SubItems(2) & " / "
                                End If
                          Next w
                                
                          If Cabecera = "" Then Cabecera = "///"
                          rankingC.DiscardSavedData
                          rankingC.Database.SetDataSource RsReporte
                          rankingC.Text20.SetText sRazonSocial
                          rankingC.Text10.SetText "T. Prod:" & IIf(Me.chkTipo.value = 0, Me.cboTipoProducto.Text, "Todos") & " /T. Ped: " & IIf(Me.chkTipoPedido.value = 0, Me.cboTipoPedido.Text, "Todos") & " /Grupo: " & IIf(Me.chkGrupo.value = 0, Me.cboGrupo.Text, "Todos") & " / SubGrupo:" & IIf(Me.chkSubGrupo.value = 0, Me.cboSubGrupo.Text, "Todos")
                          rankingC.Text9.SetText "Datos de Local: " & IIf(Cabecera <> "", Mid(Cabecera, 1, Len(Cabecera) - 2), "d")
                          rankingC.Text6.SetText "Ranking Detallado de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", "Combinaciones"))))) & ". Considerando el : " & IIf(optValor(0).value = True, optValor(0).Caption, IIf(optValor(1).value = True, optValor(1).Caption, optValor(2).Caption))
                          rankingC.Text7.SetText "Intervalo Analizado del : " & Format(Me.dtpFecIni.value, "dd/MM/yyyy") & " al " & Format(Me.dtpFecFin.value, "dd/MM/yyyy")
                          frmEmite.CRViewer.ReportSource = rankingC
        End If
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
End Sub

Private Sub chkTodo_Click()
        activaCheck IIf(chkTodo.value = 1, True, False)
End Sub

