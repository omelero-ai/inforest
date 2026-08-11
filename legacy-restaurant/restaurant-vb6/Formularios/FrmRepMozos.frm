VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRepMozos 
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7500
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
      Left            =   2145
      Picture         =   "FrmRepMozos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4845
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
      Left            =   3405
      Picture         =   "FrmRepMozos.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4845
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
      Left            =   6090
      Picture         =   "FrmRepMozos.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4845
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
      Height          =   4755
      Left            =   30
      TabIndex        =   1
      Top             =   -30
      Width           =   7410
      Begin VB.CheckBox chkMozo 
         Caption         =   "Todos los Mozos"
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
         TabIndex        =   23
         Top             =   1980
         Value           =   1  'Checked
         Width           =   1905
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
         TabIndex        =   7
         Top             =   375
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
         TabIndex        =   6
         Top             =   2385
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1830
         TabIndex        =   5
         Top             =   4125
         Value           =   -1  'True
         Width           =   1545
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
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Pedidos"
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
         Top             =   795
         Value           =   1  'Checked
         Width           =   1905
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
         TabIndex        =   2
         Top             =   1590
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1860
         TabIndex        =   8
         Top             =   3495
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641537
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1860
         TabIndex        =   9
         Top             =   3090
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641537
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   5130
         TabIndex        =   10
         Top             =   3045
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   24641539
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5130
         TabIndex        =   11
         Top             =   3465
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   24641539
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboTipoProducto 
         Height          =   345
         Left            =   1845
         TabIndex        =   25
         Top             =   360
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboTurno 
         Height          =   345
         Left            =   1845
         TabIndex        =   26
         Top             =   2370
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboGrupo 
         Height          =   345
         Left            =   1845
         TabIndex        =   27
         Top             =   1170
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboTipoPedido 
         Height          =   345
         Left            =   1845
         TabIndex        =   28
         Top             =   765
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboProducto 
         Height          =   345
         Left            =   1860
         TabIndex        =   29
         Top             =   1560
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboMozo 
         Height          =   345
         Left            =   1845
         TabIndex        =   30
         Top             =   1950
         Width           =   3135
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadFont3D      =   3
         ForeColorEven   =   -2147483640
         BackColorEven   =   -2147483624
         BackColorOdd    =   -2147483624
         RowHeight       =   423
         ExtraHeight     =   79
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5530
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mozo :"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   1965
         Width           =   540
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
         TabIndex        =   19
         Top             =   390
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
         Left            =   1170
         TabIndex        =   18
         Top             =   2340
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
         Left            =   555
         TabIndex        =   17
         Top             =   3180
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
         Left            =   675
         TabIndex        =   16
         Top             =   3570
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Reporte :"
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
         Index           =   3
         Left            =   555
         TabIndex        =   15
         Top             =   4110
         Width           =   1155
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
         Left            =   1125
         TabIndex        =   14
         Top             =   1230
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pedido :"
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
         TabIndex        =   13
         Top             =   780
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
         Left            =   885
         TabIndex        =   12
         Top             =   1620
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
      Left            =   4740
      Picture         =   "FrmRepMozos.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4845
      Width           =   1350
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   45
      Top             =   4905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRepMozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sNombre As String

Dim RsReporte As Recordset

Dim RsTempoE As Recordset
Dim RsTempoT As Recordset
Dim RsTempoC As Recordset
Dim RsTempoP As Recordset
Dim RsTurno As Recordset
'Dim RsLocal As Recordset
'Dim RsSalon As Recordset
'Dim RsSubGrupo As Recordset
'Combos
Dim RsTipoPedido As Recordset
Dim RsTipo As Recordset
Dim RsGrupo As Recordset
Dim RsProducto As Recordset
'Reportes
Dim Detallado As New DsrMozoProducto
'Cadena SQL
Dim sCriterio As String
Dim sTitulo As String

Sub LlenaCombos()
'    With cboLocal
'         Isql = "Select * from vLocal"
'         Call Lib.ConfCombo(2, cboLocal, "Cod", 2, "Codigo", 600, 0, _
'                                         "Local", 2, "Descripcion", 2500, 0)
'
'         Set RsLocal = Lib.OpenRecordset(Isql, Cn)
'         Set .DataSourceList = RsLocal
'         .DataFieldToDisplay = "Descripcion"
'         .DataFieldList = "Codigo"
'    End With

'    With cboSalon
'         Isql = "Select * from vSalon"
'         Call Lib.ConfCombo(2, cboSalon, "Cod", 2, "Codigo", 600, 0, _
'                                         "Salon", 2, "Descripcion", 2500, 0)
'
'         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
'         Set .DataSourceList = RsSalon
'         .DataFieldToDisplay = "Descripcion"
'         .DataFieldList = "Codigo"
'    End With
    With cboTipoProducto
         Isql = "Select * from vTipoProducto"
         Call Lib.ConfCombo(2, cboTipoProducto, "Cod", 2, "Codigo", 600, 0, _
                                         "Tipo Producto", 2, "Descripcion", 2500, 0)
                          
         Set RsTipo = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsTipo
         .DataFieldToDisplay = "Descripcion"
         .DataFieldList = "Codigo"
    End With
    
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Call Lib.ConfCombo(2, cboTipoPedido, "Cod", 2, "Codigo", 600, 0, _
                                              "Tipo Pedido", 2, "Descripcion", 2500, 0)
                          
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsTipoPedido
         .DataFieldToDisplay = "Descripcion"
         .DataFieldList = "Codigo"
    End With
    
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Call Lib.ConfCombo(2, cboGrupo, "Cod", 2, "Codigo", 600, 0, _
                                         "Grupo", 2, "Descripcion", 2500, 0)
                          
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsGrupo
         .DataFieldToDisplay = "Descripcion"
         .DataFieldList = "Codigo"
    End With
    
    With cboProducto
         Isql = "Select * from vProducto order by descripcion"
         Call Lib.ConfCombo(2, cboProducto, "Cod", 2, "Codigo", 800, 0, _
                                            "Producto", 2, "Descripcion", 3100, 0)
                          
         Set RsProducto = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsProducto
         .DataFieldToDisplay = "Descripcion"
         .DataFieldList = "Codigo"
    End With
    
    With cboMozo
         Isql = "Select * from vMozo"
         Call Lib.ConfCombo(2, cboMozo, "Cod", 2, "Codigo", 600, 0, _
                                         "Mozo", 2, "Descripcion", 2500, 0)
                          
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsMozo
         .DataFieldToDisplay = "Descripcion"
         .DataFieldList = "Codigo"
    End With

    
    With cboTurno
         Isql = "Select * from MTURNO order by tTurno DESC"
         Call Lib.ConfCombo(4, cboTurno, "Turno", 2, "tTurno", 1200, 0, _
                                         "Caja", 2, "tCaja", 750, 0, _
                                         "Fec.Ini", 2, "fInicial", 1750, 0, _
                                         "Fec.Fin", 2, "fFinal", 1750, 0)
                          
         Set RsTurno = Lib.OpenRecordset(Isql, Cn)
         Set .DataSourceList = RsTurno
         .DataFieldToDisplay = "tTurno"
         .DataFieldList = "tTurno"
    End With
    
'    With cboSubGrupo
'         Isql = "Select * from vSubGrupo order by Descripcion"
'         Call Lib.ConfCombo(2, cboSubGrupo, "Cod", 2, "Codigo", 600, 0, _
'                                         "SubGrupo", 2, "Descripcion", 2500, 0)
'
'         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
'         Set .DataSourceList = RsSubGrupo
'         .DataFieldToDisplay = "Descripcion"
'         .DataFieldList = "Codigo"
'    End With
    
End Sub

Private Sub cboGrupo_Click()
   RsProducto.Filter = "tGrupo='" & cboGrupo.Value & "'"
   cboProducto.Value = ""
End Sub

Private Sub chkMozo_Click()
If chkMozo.Value = 1 Then
   cboMozo.Enabled = False
   cboMozo.Text = ""
Else
   cboMozo.Enabled = True
End If
End Sub

Private Sub chkProducto_Click()
If chkProducto.Value = 1 Then
   cboProducto.Enabled = False
   cboProducto.Text = ""
Else
   cboProducto.Enabled = True
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
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
      RsProducto.Filter = adFilterNone
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

Private Sub chkLocal_Click()
   If chkLocal.Value = 1 Then
      cboLocal.Enabled = False
      cboLocal.Text = ""
      RsSalon.Filter = adFilterNone
   Else
      cboLocal.Enabled = True
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

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
   
   If chkTurno.Value = 0 Then
      If cboTurno.Text = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
      RsTurno.MoveFirst
      RsTurno.Find ("tTurno='" & cboTurno.Value & "'")
      sTitulo = " Del " & Format(RsTurno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(RsTurno!fFinal), Format(RsTurno!finicial, "dd/mmm/yyyy") & " 23:59", RsTurno!fFinal), "dd/mmm/yyyy HH:mm")
      
      sCriterio = " tTurno ='" & cboTurno.Value & "'"
      
   Else
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
      sCriterio = sCriterio & " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
      
   End If
   
   'If chkSalon.Value = 0 Then
   '   If cboSalon.Text = "" Then
   '      MsgBox "Debe escoger un Salón", vbCritical, sMensaje
   '      Exit Sub
   '   End If
   '   sCriterio = sCriterio & "and tSalon ='" & cboSalon.Value & "'"
   'End If
   
   If chkTipo.Value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de Producto", vbCritical, sMensaje
         Exit Sub
      End If
      ''''''''''''''
      sCriterio = sCriterio & " and vProducto.TipoProducto ='" & cboTipoProducto.Text & "'"
   End If
   
   If chkMozo.Value = 0 Then
      If cboMozo.Text = "" Then
         MsgBox "Debe escoger un Mozo", vbCritical, sMensaje
        Exit Sub
      End If
      ''''''''''''''''''''
      sCriterio = sCriterio & " and MPEDIDO.tMozo ='" & cboMozo.Value & "'"
   End If
   
   If chkTipoPedido.Value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.Value & "'"
   End If
      
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.Grupo='" & cboGrupo.Text & "'"
   End If
   
   If chkProducto.Value = 0 Then
      If cboProducto.Text = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.Descripcion ='" & cboProducto.Text & "'"
   End If
   
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               'frmEmite.CRViewer.DisplayGroupTree = False
               'Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               'Detallado.PaperSize = crPaperLetter
               'Detallado.PaperOrientation = crLandscape
               'frmEmite.CRViewer.ViewReport
               'frmEmite.Show vbModal
               'Me.MousePointer = vbDefault
               
               If optOpcion(0).Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = True
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
               
          Case Is = 1 ' Imprimir
               Genera
               If RsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
               
               Screen.MousePointer = vbDefault
               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PrintOut
               End If
               
          Case Is = 2 ' Salir
               Unload Me
          
          Case Is = 3 ' Exportar
               Genera
               If RsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                
               Screen.MousePointer = vbDefault
               If optOpcion(0).Value = True Then
                  Detallado.ExportOptions.FormatType = crEFTExcel50
                  Detallado.ExportOptions.DestinationType = crEDTDiskFile
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado.Export False
               End If
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub chkTurno_Click()
   If chkTurno.Value = 1 Then
      cboTurno.Enabled = False
      cboTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
   Else
      cboTurno.Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    Me.Caption = "Producción por Mozos"
    LlenaCombos
    
    sNombre = ""
    sNombre = dbTemporal(sCaja, 8, "Mozo", "nVarChar(30)", "Producto", "nVarChar(50)", "Cantidad", "Float", "Venta", "Float", "Monto", "Float", "TipoProducto", "nVarChar(30)", "Grupo", "nVarChar(50)", "tTipoPedido", "nVarchar(2)")
    
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    
    cboTurno.Enabled = False
    cboTurno.Text = ""
    
    cboTipoProducto.Enabled = False
    cboTipoProducto.Text = ""
    
    cboTipoPedido.Enabled = False
    cboProducto.Enabled = False
    
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    
    cboMozo.Enabled = False
    cboMozo.Text = ""
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
   Set RsLocal = Nothing
   Set RsMozo = Nothing
   
   Cn.Execute "Delete from TTABLATEMP where Tabla='" & Trim(sNombre) & "'"
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
   
   Dim sOrden As String
   Screen.MousePointer = vbHourglass
   
   Cn.Execute "Delete from " & sNombre
   
   Isql = "Insert into " & sNombre & " (Mozo,Producto,Cantidad,Venta,Monto,TipoProducto,Grupo,tTipoPedido ) " _
        & " SELECT dbo.vMozo.tResumido AS MOZO, dbo.vProducto.Descripcion AS Producto," _
        & " dbo.DPEDIDO.nCantidad AS Cantidad, dbo.DPEDIDO.nPrecioVenta AS nVenta,DPEDIDO.nCantidad * DPEDIDO.nPrecioVenta As Monto, dbo.vProducto.TipoProducto, dbo.vProducto.Grupo,dbo.DPEDIDO.tTipoPedido" _
        & " FROM dbo.MPEDIDO LEFT OUTER JOIN " _
        & " dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo RIGHT OUTER JOIN " _
        & " dbo.DPEDIDO LEFT OUTER JOIN " _
        & " dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido" _
        & " WHERE MPEDIDO.tEstadoPedido<>'03' And DPEDIDO.tEstadoItem='N' AND " & sCriterio
   
   '    & " SELECT dbo.vProducto.TipoProducto, dbo.vProducto.Grupo, dbo.vProducto.Descripcion AS Producto, dbo.DPEDIDO.nCantidad AS Cantidad, " _
        & " dbo.DPEDIDO.nPrecioVenta AS nVenta, dbo.DPEDIDO.tTipoPedido " _
        & " FROM dbo.MPEDIDO RIGHT OUTER JOIN " _
        & " dbo.DPEDIDO LEFT OUTER JOIN " _
        & " dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido"

   Cn.Execute Isql
   
   If optOpcion(0).Value = True Then
      Isql = "SELECT Mozo,Producto,SUM(Cantidad) AS Cantidad, SUM(Venta) AS Ventas,SUM(Monto)As Monto,TipoProducto,Grupo " & _
             "From " & sNombre & " Group By Mozo,TipoProducto,Grupo,Producto"
      Set RsReporte = Lib.OpenRecordset(Isql, Cn)
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource RsReporte
      Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & cboTurno.Text) & sTitulo
      frmEmite.CRViewer.ReportSource = Detallado
   End If
         
End Sub


