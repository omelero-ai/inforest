VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepVentaCompAnual 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta Comparativa Anual"
   ClientHeight    =   8925
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepVentaCompAnual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   6525
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
      Left            =   3285
      Picture         =   "frmRepVentaCompAnual.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8250
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
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
      Height          =   8175
      Left            =   0
      TabIndex        =   15
      Top             =   30
      Width           =   6510
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
         Height          =   1245
         Left            =   3360
         TabIndex        =   24
         Top             =   5895
         Width           =   2880
         Begin VB.OptionButton optOrigen 
            Caption         =   "Ventas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   25
            Top             =   585
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkCFacturados 
         Caption         =   "Valorizar con Cero la Pre Venta Facturada"
         Height          =   405
         Left            =   1200
         TabIndex        =   14
         Top             =   7725
         Width           =   4155
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
         Height          =   1305
         Left            =   495
         TabIndex        =   22
         Top             =   5895
         Width           =   2640
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   585
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   27
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.OptionButton optAgrupacion 
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   19
         Top             =   4905
         Width           =   225
      End
      Begin VB.OptionButton optAgrupacion 
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   1740
         Width           =   225
      End
      Begin VB.Frame fraproducto 
         Caption         =   " Agrupado por Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   510
         TabIndex        =   9
         Top             =   4905
         Width           =   5865
         Begin VB.TextBox txtTipo 
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
            Left            =   570
            TabIndex        =   21
            Top             =   330
            Width           =   2895
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
            Left            =   3570
            Picture         =   "frmRepVentaCompAnual.frx":082E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   270
            Width           =   765
         End
         Begin VB.TextBox txtCodProducto 
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   330
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chkproducto 
            Caption         =   "Productos"
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
            Left            =   4440
            TabIndex        =   11
            Top             =   300
            Value           =   1  'Checked
            Width           =   1260
         End
      End
      Begin VB.Frame frasubGrupo 
         Caption         =   " Agrupado por Sub Grupos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   510
         TabIndex        =   8
         Top             =   1740
         Width           =   5865
         Begin VB.OptionButton OptSel 
            Caption         =   "Ninguno"
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
            Height          =   255
            Index           =   1
            Left            =   4230
            TabIndex        =   32
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton OptSel 
            Caption         =   "Todos"
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
            Height          =   240
            Index           =   0
            Left            =   3150
            TabIndex        =   31
            Top             =   300
            Width           =   1095
         End
         Begin TrueOleDBGrid80.TDBGrid grdSubGrupos 
            Height          =   2205
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3889
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   3
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
            MultiSelect     =   0
            DeadAreaBackColor=   16777215
            ScrollTrack     =   -1  'True
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.fgcolor=&H8B1512&"
            _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.bgcolor=&H8CFFFF&"
            _StyleDefs(22)  =   "Splits(0).Style:id=67,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=88,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=69,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=71,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=87,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=89,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=90,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=94,.parent=67"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=68,.alignment=0"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=69"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=71"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0"
            _StyleDefs(40)  =   "Named:id=34:Heading"
            _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=34,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=35:Footing"
            _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=36:Selected"
            _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=37:Caption"
            _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(49)  =   "Named:id=38:HighlightRow"
            _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(51)  =   "Named:id=39:EvenRow"
            _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=40:OddRow"
            _StyleDefs(54)  =   ":id=40,.parent=33"
            _StyleDefs(55)  =   "Named:id=41:RecordSelector"
            _StyleDefs(56)  =   ":id=41,.parent=34"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccionar :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.OptionButton optAgrupacion 
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.Frame fraGrupo 
         Caption         =   " Agrupado por Grupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   510
         TabIndex        =   5
         Top             =   780
         Width           =   5865
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
            Left            =   3450
            TabIndex        =   7
            Top             =   315
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin MSDataListLib.DataCombo cboGrupo 
            Height          =   315
            Left            =   180
            TabIndex        =   6
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
      End
      Begin MSComCtl2.DTPicker dtpAnual 
         Height          =   405
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   16384003
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   375
         Left            =   3780
         TabIndex        =   13
         Top             =   7305
         Width           =   1485
         _ExtentX        =   2619
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
         CustomFormat    =   "HH"
         Format          =   16384003
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Cierre y Apertura :"
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   23
         Top             =   7365
         Width           =   2370
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   16
         Top             =   322
         Width           =   450
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
      Left            =   4740
      Picture         =   "frmRepVentaCompAnual.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8250
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
      Left            =   1830
      Picture         =   "frmRepVentaCompAnual.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8250
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
      Left            =   360
      Picture         =   "frmRepVentaCompAnual.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   135
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepVentaCompAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDetalle As String
Dim sGrafico As String
Dim ReporteAnual As New dsrVentaCompAnual
Dim RsTemporal As Recordset
Dim RsLocal As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim rsReporte As Recordset
Dim sCriterio As String
Dim sCriterio2 As String
Dim sPrecio As String
Dim sTexto As String

'CESAR 0020-2013
Dim sOrigen As String
Dim sTipo As String

'pierre
Dim VentAnu_RsSubGrupos As Recordset
Dim sSubGrupos As String
Dim colu As TrueOleDBGrid80.Columns

Sub LlenaCombos()
Dim i As Integer
    With cboGrupo
         Isql = "Select * from vGrupo where lActivo=1 order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
End Sub

Private Sub cboGrupo_Change()
'    cboSubGrupo.BoundText = ""
'    With cboSubGrupo
'         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
'         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
'         Set .RowSource = RsSubGrupo
'    End With
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
        cmdBusca.Enabled = False
        txtTipo = ""
    Else
        cmdBusca.Enabled = True
        chkGrupo.Value = 1
        cboGrupo.BoundText = ""
        'chkSubGrupo.value = 1
        'cboSubGrupo.BoundText = ""
    End If
End Sub

'Private Sub chkSubGrupo_Click()
'   If chkSubGrupo.value = 1 Then
'      cboSubGrupo.Enabled = False
'      cboSubGrupo.Text = ""
'   Else
'      cboSubGrupo.Enabled = True
'   End If
'End Sub


Sub Validar_Todos_Ninguno()
Dim RS0, RS1, RS2 As Recordset
Set RS0 = Lib.OpenRecordset("Select * from " & sSubGrupos, Cn)
'Set RS1 = Lib.OpenRecordset("Select * from " & sSubGrupos & " where filtrar=0", Cn) 'No Todos
Set RS1 = VentAnu_RsSubGrupos.Clone
'Set RS2 = Lib.OpenRecordset("Select * from " & sSubGrupos & " where filtrar=-1", Cn) 'No Ninguno
Set RS2 = VentAnu_RsSubGrupos.Clone

RS1.Filter = "Filtrar=0"
RS2.Filter = "Filtrar=-1"

'If RS1.RecordCount = 0 Then
'OptSel(0).value = True
'Else
'OptSel(0).value = False
'End If
'
'If RS2.RecordCount = 0 Then
'OptSel(1).value = True
'Else
'OptSel(1).value = False
'End If
End Sub
'--------------------------pierre

Private Sub cmdBusca_Click()
    Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lActivo=1 order by Descripcion"
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
    txtCodProducto = sCodigo
    txtTipo = sDescrip

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sCriterio2 = ""
                
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.Grupo='" & cboGrupo.Text & "'"
      sCriterio2 = "and Grupo ='" & cboGrupo.Text & "'"
   End If
   
   If optAgrupacion(1).Value = True Then
'     If OptSel(0).value = True Then
'                sCriterio = ""
'     ElseIf OptSel(1).value = True Then
'        MsgBox "Debe escoger minimo un subgrupo", vbCritical, sMensaje
'        Exit Sub
'    ElseIf OptSel(0).value = False And OptSel(1).value = False Then
'        sCriterio = " and vProducto.tsubGrupo in (select codigo from " & sSubGrupos & " where filtrar=1) "
'     End If
     
     If OptSel(0).Value = True Then
     sCriterio = ""
     
     ElseIf OptSel(1).Value = True Then
        Dim m As Integer
        Dim o As Integer
        Dim cadena As String
        cadena = "'"
        o = 0
        m = 0
        For m = 1 To grdSubGrupos.ApproxCount
            If grdSubGrupos.Columns(0).CellValue(m) = True Then
                o = o + 1
                If Len(cadena) = 1 Then
                    cadena = cadena + grdSubGrupos.Columns(1).CellValue(m) + "'"
                Else
                    cadena = cadena + ",'" + grdSubGrupos.Columns(1).CellValue(m) + "'"
                End If
            End If
        Next m
        If o = 0 Then
            MsgBox "Debe escoger minimo un subgrupo", vbCritical, sMensaje
           ' Hay_Datos = 0
            Exit Sub
        End If
 
  
        sCriterio = " and dpedido.tCodigoSubGrupo in (" + cadena + ")" 'select codigo from " & sSubGrupos & " where filtrar=1) "
 
    End If
   End If
   
   If optAgrupacion(2).Value Then
    If chkProducto.Value = 0 Then
        sCriterio = sCriterio & " and vProducto.Codigo='" & Me.txtCodProducto.Text & "'"
    End If
   End If
   
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsTemporal.RecordCount = 0 Then
                  Screen.MousePointer = vbDefault
                  'MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ReporteAnual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteAnual.PaperOrientation = crLandscape
               frmEmite.CRViewer.DisplayGroupTree = False
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
                    
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsTemporal.RecordCount = 0 Then
                  Screen.MousePointer = vbDefault
                  'MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ReporteAnual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteAnual.PaperOrientation = crLandscape
               ReporteAnual.PrintOut
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsTemporal.RecordCount = 0 Then
                  Screen.MousePointer = vbDefault
                  'MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                
               Screen.MousePointer = vbDefault
               ReporteAnual.ExportOptions.FormatType = 21
               ReporteAnual.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               ReporteAnual.ExportOptions.DiskFileName = cmdSave.FileName
               ReporteAnual.Export False
   End Select
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpAnual.Value = Date
    
    sDetalle = dbTemporal(sCaja, 39, "tGrupo1", "nVarchar(4)", "tGrupo2", "nVarchar(7)", _
                                     "C1", "Float", "M1", "Float", "Mt1", "Float", _
                                     "C2", "Float", "M2", "Float", "Mt2", "Float", _
                                     "C3", "Float", "M3", "Float", "Mt3", "Float", _
                                     "C4", "Float", "M4", "Float", "Mt4", "Float", _
                                     "C5", "Float", "M5", "Float", "Mt5", "Float", _
                                     "C6", "Float", "M6", "Float", "Mt6", "Float", _
                                     "C7", "Float", "M7", "Float", "Mt7", "Float", _
                                     "C8", "Float", "M8", "Float", "Mt8", "Float", _
                                     "C9", "Float", "M9", "Float", "Mt9", "Float", _
                                     "C10", "Float", "M10", "Float", "Mt10", "Float", _
                                     "C11", "Float", "M11", "Float", "Mt11", "Float", _
                                     "C12", "Float", "M12", "Float", "Mt12", "Float", "linea", "nVarchar(50)")
    
    LlenaCombos
    
    'Pierre ----------------------
     sSubGrupos = dbTemporal(sCaja, 5, "Filtrar", "bit", "Codigo", "nVarchar(4)", _
                                     "Descripcion", "nVarchar(50)", "Fila", "nVarchar(4)", _
                                     "Grupo", "nVarchar(50)")
    
    
    
    Cn.Execute "insert into " & sSubGrupos & "(Filtrar, codigo, descripcion, fila, grupo) select convert(bit,'1'), Codigo, Descripcion, ROW_NUMBER() OVER(ORDER BY Descripcion asc), Grupo from vsubgrupo order by descripcion"

    
    Set VentAnu_RsSubGrupos = Lib.OpenRecordset(" select Filtrar, Codigo, Descripcion, Fila, Grupo from " & sSubGrupos & " order by descripcion", Cn)

    Call ConfGrilla(5, grdSubGrupos, "Filtrar", 2, "Filtrar", 500, 2, 4, "", _
                                    "Codigo", 2, "Codigo", 0, 0, 0, "", _
                                   "Descripcion", 2, "Descripcion", 3000, 0, 0, "", _
                                   "Fila", 2, "Fila", 0, 0, 0, "", _
                                   "Grupo", 2, "Grupo", 1500, 0, 0, "")
    
    grdSubGrupos.DataSource = VentAnu_RsSubGrupos
        
    OptSel(0).Value = True
    With grdSubGrupos
        .Columns(1).Visible = False
        .Columns(3).Visible = False
    End With
    '---------------------- Pierre
    cboGrupo.Text = ""
    optAgrupacion_Click 0
End Sub

'Pierre ------------------------------------
Private Sub grdSubGrupos_FilterChange()
On Error GoTo errHandler
    Set colu = grdSubGrupos.Columns
    Dim c As Integer
    c = grdSubGrupos.col
    grdSubGrupos.HoldFields
    VentAnu_RsSubGrupos.Requery
    VentAnu_RsSubGrupos.Filter = getFilter(colu)
    grdSubGrupos.col = c
    grdSubGrupos.EditActive = True
    Exit Sub

errHandler:
    'MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(colu)
    VentAnu_RsSubGrupos.Filter = adFilterNone
End Sub

Private Sub grdSubGrupos_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 0 Then

Validar_Todos_Ninguno
End If
End Sub

'Private Sub grdSubGrupos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'    Dim columna, fila As Integer
'    Dim bookMark As Variant
'    Dim codigo As String
'    Dim falso As Boolean
'    columna = grdSubGrupos.ColContaining(X)
'    fila = grdSubGrupos.RowContaining(Y)
'    If fila >= 0 Then
'    bookMark = grdSubGrupos.RowBookmark(fila)
'
'    If bookMark <> "" Then
'        codigo = grdSubGrupos.Columns(1).CellValue(bookMark)
'        Dim filtro As String
'            If codigo <> "" Then
'            filtro = Calcular("select filtrar as codigo from " & sSubGrupos & " where codigo = '" & codigo & "'", Cn)
'            If filtro = "Verdadero" Then
'            Cn.Execute "update " & sSubGrupos & " set Filtrar=0 where codigo ='" & codigo & "' and Filtrar=1"
'            ElseIf filtro = "Falso" Then
'            Cn.Execute "update " & sSubGrupos & " set Filtrar=1 where codigo ='" & codigo & "' and Filtrar=0"
'            End If
'
'        End If
'        '    TDBGrid1.DataSource = Nothing
'        '    grdSubGrupos.DataSource = Nothing
'        'VentAnu_RsSubGrupos.Requery
'        Set VentAnu_RsSubGrupos = Lib.OpenRecordset("Select * from " & sSubGrupos & " order by descripcion", Cn)
'
'        '    TDBGrid1.DataSource = VentAnu_RsSubGrupos
'        '    grdSubGrupos.DataSource = VentAnu_RsSubGrupos
'
'        grdSubGrupos.DataSource = VentAnu_RsSubGrupos
'        '    TDBGrid1.Refresh
'        '    grdSubGrupos.Refresh
'        Validar_Todos_Ninguno
'        'grdSubGrupos.col = columna
'        grdSubGrupos.Row = fila
'        grdSubGrupos.col = columna
'        grdSubGrupos.SetFocus
'
'        Set colu = grdSubGrupos.Columns
'        Call cmdClearFilter_Click(colu)
'        VentAnu_RsSubGrupos.Filter = adFilterNone
'        End If
'    Else
'    Exit Sub
'    End If
'End If
'End Sub
'------------------------------------ Pierre

Private Sub optAgrupacion_Click(Index As Integer)
 fraGrupo.Enabled = False
 chkGrupo.Enabled = False
 cboGrupo.Enabled = False
 fraSubGrupo.Enabled = False
 fraproducto.Enabled = False
 chkProducto.Enabled = False
 OptSel(0).Enabled = False
 OptSel(1).Enabled = False
 grdSubGrupos.Enabled = False

   Select Case Index
          Case Is = 0  'Grupo
               fraGrupo.Enabled = True
               chkGrupo.Enabled = True
               cboGrupo.Enabled = True
               
          Case Is = 1  'Sub Grupo
               fraSubGrupo.Enabled = True
               OptSel(0).Enabled = True
               OptSel(1).Enabled = True
               grdSubGrupos.Enabled = True
               
          Case Is = 2  'Producto
               fraproducto.Enabled = True
               chkProducto.Enabled = True
   End Select
   
   cboGrupo.Enabled = False
   cboGrupo.Text = ""
   chkGrupo.Value = 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sDetalle
   
   Set RsTemporal = Nothing
   Set rsReporte = Nothing
   Set VentAnu_RsSubGrupos = Nothing
    Cn.Execute "drop table " & sSubGrupos
End Sub

Public Sub Genera()
    Screen.MousePointer = vbHourglass
    Dim sFecha1 As String
    Dim sFecha2 As String
    Dim sAno1 As String
    Dim sAno2 As String
    Dim sMes As String
    Dim sDia As String
    
    'CESAR 0020-2013
    If chkCFacturados.Value = 0 Then
        'VENTA
        If optOrigen(0).Value = True Then
            sOrigen = "00"
            If optValor(0).Value Then
               sTipo = "00"
               sPrecio = "dbo.DPEDIDO.nVenta"
               sTexto = " con Precios de Venta"
            ElseIf optValor(1).Value Then
               sTipo = "00"
               sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
               sTexto = " con Precios Netos"
            Else
               sTipo = "01"
               sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
               sTexto = " con Precios de Costo"
            End If
        End If
        'COMBINACION
        If optOrigen(4).Value = True Then
            sOrigen = "04"
            If optValor(0).Value Then
               sTipo = "00"
               sPrecio = "dbo.CPEDIDO.nVenta"
               sTexto = " con Precios de Venta"
            ElseIf optValor(1).Value Then
               sTipo = "00"
               sPrecio = "dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad"
               sTexto = " con Precios Netos"
            Else
               sTipo = "01"
               sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
               sTexto = " con Precios de Costo"
            End If
        End If
        
    Else
        'VENTA
        If optOrigen(0).Value = True Then
            sOrigen = "00"
            If optValor(0).Value Then
                sTipo = "00"
                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nVenta end "
                sTexto = " con Precios de Venta"
            ElseIf optValor(1).Value Then
                sTipo = "00"
                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                sTexto = " con Precios Netos"
            Else
               sTipo = "01"
               sPrecio = " case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
               sTexto = " con Precios de Costo"
            End If
        End If
        'COMBINACION
        If optOrigen(4).Value = True Then
            sOrigen = "04"
            If optValor(0).Value Then
                sTipo = "00"
                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta end "
                sTexto = " con Precios de Venta"
            ElseIf optValor(1).Value Then
                sTipo = "00"
                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad end "
                sTexto = " con Precios Netos"
            Else
               sTipo = "01"
               sPrecio = " case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad end "
               sTexto = " con Precios de Costo"
            End If
        End If
        
    End If
    
    sFecha1 = ""
    sFecha2 = ""
    sAno1 = LTrim(str(Year(dtpAnual.Value)))
    sAno2 = LTrim(str(Year(dtpAnual.Value) - 1))
        
    sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/12/31') "
    sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/12/31') "
    
    
    
    'CESAR 0020-2013
    If sOrigen = "00" Then
    
       If sTipo = "00" Then '--------VENTAS y NETO
           Isql = "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
                "union all " & _
                "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
       Else '----COSTOS ---
            Isql = "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TPRODUCTOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nGasto,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                "LEFT JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTOPROPIEDAD.tProducto AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem AND TPRODUCTOPROPIEDAD.tEnlace<>'' " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
                "union all " & _
                "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TPRODUCTOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nGasto,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                "LEFT JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTOPROPIEDAD.tProducto AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem AND TPRODUCTOPROPIEDAD.tEnlace<>'' " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
        End If
    Else
        If sTipo = "00" Then '--------VENTA y NETO DE COMBINACIONES
            Isql = "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
                "union all " & _
                "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
        Else ' ----COSTOS ----
            Isql = "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TCOMBOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nGasto,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
                "INNER JOIN dbo.tComboPropiedad ON dbo.tComboPropiedad.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.tComboPropiedad.tProducto = dbo.CPEDIDO.tProductoCombo AND dbo.tComboPropiedad.tItem = dbo.CPEDIDO.tItemCombo " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
                "union all " & _
                "(SELECT '' AS tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TCOMBOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nGasto,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
                "INNER JOIN dbo.tComboPropiedad ON dbo.tComboPropiedad.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.tComboPropiedad.tProducto = dbo.CPEDIDO.tProductoCombo AND dbo.tComboPropiedad.tItem = dbo.CPEDIDO.tItemCombo " & _
                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.Value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.Value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
        End If
    End If

    
    Set RsTemporal = Lib.OpenRecordset(Isql, Cn)
     
    If RsTemporal.RecordCount = 0 Then
      MsgBox "No hay Datos para Mostrar", vbInformation, sMensaje
      Exit Sub
    End If
    Cn.Execute "delete from " & sDetalle
            
    Dim sLinea As String
    Dim i As Integer
    
    RsTemporal.MoveFirst
    Do While Not RsTemporal.EOF
                 
       If Not IsNull(RsTemporal!Fecha) Then
       
          If optAgrupacion(0).Value = True Then
             If Calcular("select count(tGrupo1) as Codigo from " & sDetalle & " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "'", Cn) > 0 Then
                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value) - 1, 0, RsTemporal!nVenta) & _
                       " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "'"
             Else
                sLinea = ""
               
                For i = 1 To 12
                    If RsTemporal!Fecha <> i Then
                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
                    End If
                Next i
                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
                       "values ('" & IIf(IsNull(RsTemporal!Grupo), "Sin Grupo...", RsTemporal!Grupo) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.Value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.Value), RsTemporal!nVenta, 0) & " )"
             End If
                                 
          ElseIf optAgrupacion(1).Value = True Then
                    
             If Calcular("select count(tGrupo1) as Codigo from " & sDetalle & " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "'", Cn) > 0 Then
                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value) - 1, 0, RsTemporal!nVenta) & _
                       " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "'"
             Else
                sLinea = ""
                For i = 1 To 12
                    If RsTemporal!Fecha <> i Then
                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
                    End If
                Next i
                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
                       "values ('" & IIf(IsNull(RsTemporal!SubGrupo), "Sin Sub-Grupo...", RsTemporal!SubGrupo) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.Value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.Value), RsTemporal!nVenta, 0) & " )"
             End If
          Else
                
             If Calcular("select count(tGrupo2) as Codigo from " & sDetalle & " where tGrupo2 ='" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "'", Cn) > 0 Then
                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.Value) - 1, 0, RsTemporal!nVenta) & _
                       " where tGrupo2 ='" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "'"
             Else
                sLinea = ""
                
                
                For i = 1 To 12
                    If RsTemporal!Fecha <> i Then
                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
                    End If
                Next i
                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
                       "values ('" & IIf(IsNull(RsTemporal!nomprod), "Sin Nombre...", RsTemporal!nomprod) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.Value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.Value), RsTemporal!nVenta, 0) & " )"
             End If
          End If
          Cn.Execute Isql
       End If
       RsTemporal.MoveNext
    Loop
        Isql = "select * from " & sDetalle
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
    Dim sMeses(13) As String
    Dim nMonto As Double
    
    ReporteAnual.DiscardSavedData
    ReporteAnual.Database.SetDataSource rsReporte
    ReporteAnual.PaperOrientation = crLandscape
    
    ReporteAnual.ReportTitle = " Comparativo del  Año 1: " & dtpAnual.Year - 1 & " / Año 2 : " & dtpAnual.Year & sTexto
    ReporteAnual.Text20.SetText sRazonSocial
    ReporteAnual.Text6.SetText localConectado
    frmEmite.CRViewer.ReportSource = ReporteAnual
End Sub

Private Sub optTipo_Click(Index As Integer)
   If Index = 0 Then
      fraGrupo.Enabled = True
      fraSubGrupo.Enabled = False
      fraproducto.Enabled = False
      optAgrupacion(0).Enabled = True
      optAgrupacion(1).Enabled = True
      optAgrupacion(2).Enabled = True
      optAgrupacion(0).Value = True
      optAgrupacion(1).Value = False
      optAgrupacion(2).Value = False
      chkGrupo.Enabled = True
      'chkSubGrupo.Enabled = False
      chkProducto.Enabled = False
      cboGrupo.Enabled = True
      cboGrupo.Text = ""
      'cboSubGrupo.Enabled = False
      'cboSubGrupo.Text = ""
      txtTipo.Enabled = False
      txtTipo.Text = ""
   Else
      fraproducto.Enabled = False
      fraGrupo.Enabled = False
      fraSubGrupo.Enabled = False
      chkProducto.Enabled = False
      chkGrupo.Enabled = False
      'chkSubGrupo.Enabled = False
      optAgrupacion(0).Enabled = False
      optAgrupacion(1).Enabled = False
      optAgrupacion(2).Enabled = False
      txtTipo.Enabled = False
      txtTipo.Text = ""
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
      'cboSubGrupo.Enabled = False
      'cboSubGrupo.Text = ""
   End If
End Sub

Private Sub OptSel_Click(Index As Integer)
If Index = 0 Then
    Cn.Execute "update " & sSubGrupos & " set Filtrar=-1"
    
ElseIf Index = 1 Then
    Cn.Execute "update " & sSubGrupos & " set Filtrar=0"
End If

   'VentAnu_RsSubGrupos.Requery
   Set VentAnu_RsSubGrupos = Lib.OpenRecordset(" select Filtrar, Codigo, Descripcion, Fila, Grupo from " & sSubGrupos & " order by descripcion", Cn)
   VentAnu_RsSubGrupos.Filter = adFilterNone
   grdSubGrupos.DataSource = VentAnu_RsSubGrupos
End Sub
