VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepControlEnviosAutorizados 
   Caption         =   "Control de Envios Autorizados"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
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
      Left            =   3555
      Picture         =   "frmRepControlEnviosAutorizados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5040
      Visible         =   0   'False
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
      Left            =   4905
      Picture         =   "frmRepControlEnviosAutorizados.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5040
      Width           =   1350
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
      Left            =   2220
      Picture         =   "frmRepControlEnviosAutorizados.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5040
      Width           =   1350
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
      Left            =   960
      Picture         =   "frmRepControlEnviosAutorizados.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7395
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
         TabIndex        =   28
         Top             =   2640
         Width           =   7080
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
            TabIndex        =   32
            Top             =   270
            Width           =   1125
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
            TabIndex        =   31
            Top             =   660
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CheckBox chkMozo 
            Caption         =   "Todos los Meseros"
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
            Left            =   4560
            TabIndex        =   30
            Top             =   240
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.CheckBox chkUsuario 
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
            Left            =   4560
            TabIndex        =   29
            Top             =   600
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin MSDataListLib.DataCombo cboMozo 
            Height          =   315
            Left            =   1300
            TabIndex        =   33
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
         Begin MSDataListLib.DataCombo cboUsuario 
            Height          =   315
            Left            =   1300
            TabIndex        =   34
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
         Left            =   4680
         TabIndex        =   20
         Top             =   2265
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
         Left            =   4680
         TabIndex        =   19
         Top             =   735
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
         Left            =   4680
         TabIndex        =   18
         Top             =   1125
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
         Left            =   4680
         TabIndex        =   17
         Top             =   1500
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
         Index           =   0
         Left            =   3780
         Picture         =   "frmRepControlEnviosAutorizados.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1455
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
         Left            =   3780
         Picture         =   "frmRepControlEnviosAutorizados.frx":147A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2220
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1485
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2250
         Width           =   2265
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1875
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
         Index           =   2
         Left            =   3780
         Picture         =   "frmRepControlEnviosAutorizados.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1845
         Width           =   765
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1890
         Value           =   1  'Checked
         Width           =   2115
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
         Left            =   4680
         TabIndex        =   7
         Top             =   375
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   4230
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   3810
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   330
         Left            =   3630
         TabIndex        =   3
         Top             =   3810
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   330
         Left            =   3630
         TabIndex        =   4
         Top             =   4230
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   5400
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1395
         TabIndex        =   8
         Top             =   360
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
         Left            =   1410
         TabIndex        =   21
         Top             =   720
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
         Left            =   1410
         TabIndex        =   22
         Top             =   1110
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
         Left            =   705
         TabIndex        =   27
         Top             =   2310
         Width           =   630
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
         Left            =   690
         TabIndex        =   26
         Top             =   780
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
         Left            =   360
         TabIndex        =   25
         Top             =   1170
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
         Left            =   435
         TabIndex        =   24
         Top             =   1545
         Width           =   900
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
         Left            =   615
         TabIndex        =   23
         Top             =   1935
         Width           =   720
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
         Left            =   720
         TabIndex        =   9
         Top             =   420
         Width           =   615
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
         Left            =   195
         TabIndex        =   6
         Top             =   4290
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
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   3840
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmRepControlEnviosAutorizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset

Dim RsSalon As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsMozo As Recordset
Dim RsUsuario As Recordset
Dim Rsturno As Recordset

Dim sProducto As String
Dim sTurno As String
Dim sCliente As String
Dim fInicio As Date
Dim fFinal As Date

Dim sTitulo As String
Dim sCriterio As String
Dim xCriterio As String

Dim Detallado As New dsrControlEnviosAutorizados

Dim tMozoCajero As String


Sub LlenaCombos()
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
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
    
    With cboMozo
         Isql = "Select * from vMozo where substring(Codigo,1,1)<>'*' and lAutorizaEnviosMozo=1"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUsuario
         Isql = "Select * from TUSUARIO where substring(tCodigoUsuario,1,1)<>'*'"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "tDetallado"
             .ListField = "tDetallado"
             .BoundColumn = "tCodigoUsuario"
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
            Isql = "SELECT dbo.TDELIVERY.tCodigoDelivery AS Codigo, dbo.TDELIVERY.tTelefono, LTRIM(dbo.TDELIVERY.tApellido) + ' ' + LTRIM(dbo.TDELIVERY.tNombre) AS Descripcion From dbo.TDELIVERY Where (dbo.TDELIVERY.lActivo = 1)"
            frmBusca.cboCriterio.Enabled = True
            frmBusca.nPredeterm = 2
              Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                                "Identificador", 2, "tTelefono", 1500, 0, 0, "", _
                                                "Cliente", 2, "Descripcion", 4500, 0, 0, "")
            frmBusca.Show vbModal
            If Not wEnter Then
               Exit Sub
            End If
            sCliente = sCodigo
            txtCliente.Text = sDescrip
   End Select

End Sub

Private Sub chkSalon_Click()
   If chkSalon.Value = 1 Then
      cboSalon.Text = ""
      cboSalon.Enabled = False
   Else
      cboSalon.Enabled = True
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

Private Sub chkCliente_Click()
   If chkCliente.Value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
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
            
        Else
            MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
        End If
   Else
        If dtpFecIni.Value + dtpHoraIni.Value > dtpFecFin.Value + dtpHoraFin.Value Then
            MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
            Exit Sub
            dtpFecFin.SetFocus
        End If
        sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHoraIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHoraFin.Value, "HH:mm")
    
    End If
    
    If chkSalon.Value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salon", vbCritical, sMensaje
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

            Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Detallado.PaperOrientation = crPortrait

            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
               
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
                
            Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Detallado.PaperOrientation = crPortrait
            Detallado.PrintOut

        Case Is = 2 ' Salir
            Unload Me
          
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If

            Detallado.ExportOptions.FormatType = 21
            Detallado.ExportOptions.DestinationType = 1
            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
            cmdSave.ShowSave
            If cmdSave.FileName = "" Then
                 Exit Sub
            End If
            Detallado.ExportOptions.DiskFileName = cmdSave.FileName
            Detallado.Export False

   End Select
    
End Sub

Public Sub Genera()
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_ControlEnviosAutorizados", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
        
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/DD") & " " & Format(dtpHoraFin.Value, "HH:mm")
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/DD") & " " & Format(dtpHoraIni.Value, "HH:mm")
        
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 30, cboSalon.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@tCodigoCliente", adVarChar, adParamInput, 30, sCliente
    oComando.CreateParameter "@turno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@ttipodato", adVarChar, adParamInput, 2, tMozoCajero
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 30, cboMozo.BoundText
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 50, cboUsuario.Text
    
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    
      Detallado.DiscardSavedData
      'Detallado.Text8.SetText localConectado
      Detallado.Database.SetDataSource rsReporte
      'Detallado.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Local/Salon"
      Detallado.Text10.SetText sRazonSocial
      Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
      'Detallado.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = Detallado
      
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
      dtpHoraIni.Enabled = True
      dtpHoraFin.Enabled = True
   Else
      cmdBusca(1).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHoraIni.Enabled = False
      dtpHoraFin.Enabled = False
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

Private Sub chkUsuario_Click()
    If chkUsuario.Value = 1 Then
       cboUsuario.Enabled = False
       cboUsuario.Text = ""
    Else
       cboUsuario.Enabled = True
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

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    
    Isql = "Select * from MTURNO order by tTurno DESC"
    Set Rsturno = Lib.OpenRecordset(Isql, Cn)
    
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    cmdBusca(2).Enabled = False
    
    tMozoCajero = "U"
    
    cboSalon.Enabled = False
    cboSalon.Text = ""
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    cboMozo.Enabled = False
    cboMozo.Text = ""
    cboUsuario.Enabled = False
    cboUsuario.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsGrupo = Nothing
   Set RsSubgrupo = Nothing
   Set RsMozo = Nothing
   Set RsUsuario = Nothing
End Sub
