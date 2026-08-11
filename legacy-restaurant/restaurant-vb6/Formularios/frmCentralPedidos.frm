VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCentralPedidos 
   Caption         =   "Correlativo Pedidos Centro Produccion"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   13350
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   13290
      TabIndex        =   2
      Top             =   7140
      Width           =   13350
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ver Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   1
         Left            =   6890
         Picture         =   "frmCentralPedidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ver Documento"
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdEmite 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6890
         Picture         =   "frmCentralPedidos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Emite"
         Top             =   720
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpcion 
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
         Height          =   630
         Index           =   6
         Left            =   5880
         Picture         =   "frmCentralPedidos.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ver Detalle"
         Top             =   720
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Revertir Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   5
         Left            =   10200
         Picture         =   "frmCentralPedidos.frx":0A3C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Revertir Entrega"
         Top             =   30
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Confirma Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   3
         Left            =   9200
         Picture         =   "frmCentralPedidos.frx":0FC6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Confirma Entrega"
         Top             =   30
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Mod. Fecha Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   2
         Left            =   7990
         Picture         =   "frmCentralPedidos.frx":10C8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Modificar Fecha Entrega"
         Top             =   30
         Width           =   1200
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   11
         Top             =   30
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCentralPedidos.frx":150A
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmCentralPedidos.frx":1A4C
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCentralPedidos.frx":1F8E
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmCentralPedidos.frx":24D0
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmCentralPedidos.frx":2A12
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmCentralPedidos.frx":2F54
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Caption         =   "Registro"
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
            Left            =   1650
            TabIndex        =   18
            Top             =   150
            Width           =   2505
         End
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   3855
         Picture         =   "frmCentralPedidos.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filtrar"
         Top             =   700
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas de Entrega "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   165
         TabIndex        =   5
         Top             =   650
         Width           =   3660
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2205
            TabIndex        =   6
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   61865985
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   7
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   61865985
            CurrentDate     =   37539
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Del "
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
            Left            =   0
            TabIndex        =   9
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " al "
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
            Left            =   1845
            TabIndex        =   8
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ver Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   5880
         Picture         =   "frmCentralPedidos.frx":3598
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ver Detalle"
         Top             =   30
         Width           =   1000
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
         Height          =   630
         Index           =   4
         Left            =   12220
         Picture         =   "frmCentralPedidos.frx":369A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1000
      End
      Begin VB.Label txtFechaMod 
         Caption         =   "Label1"
         Height          =   255
         Left            =   8520
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6270
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   11060
         _LayoutType     =   4
         _RowHeight      =   21
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).FooterText=   "Footer Text"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
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
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
         RowDividerColor =   12632256
         RowSubDividerColor=   -2147483648
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16711935
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   225
         Top             =   2115
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmCentralPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsCabecera As Recordset
Dim nColumna As Integer
Dim nCambioFecha As String
Dim RsDetalleR As Recordset

Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Dim i As Integer

Private Sub cmdEmite_Click()
     frmRepEntrega.Show vbModal
End Sub


Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdGrilla
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
         Case Is = 0 'Detalle
                If RsCabecera.RecordCount > 0 Then
                   sPedido = grdGrilla.Columns(0).Text
                   frmPedido.cmdOpcion(5).Enabled = False
                   frmPedido.cmdOpcion(6).Enabled = False
                   frmPedido.cmdOpcion(7).Enabled = False
                   frmPedido.Show
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
                
         Case Is = 1 'Ver Pagos
                If RsCabecera.RecordCount > 0 Then
                   sPedido = grdGrilla.Columns(0).Text
                   frmPedidosVerPagos.Show vbModal
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
         Case Is = 2 ' Modificar fecha entrega
                If RsCabecera.RecordCount > 0 Then
                     'Validar si ya se entrego
                     If grdGrilla.Columns(6).Text = "ENTREGADO" Then
                          MsgBox "El Pedido ya se encuentra Entregado", vbInformation, sMensaje
                          Exit Sub
                     End If
'                    sPedido = grdGrilla.Columns(0).Text
'                    Dim lxEntregado As Boolean
'                    lxEntregado = Calcular("Select ISNULL(lEntregado,0) As codigo From MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
'                    If lxEntregado = True Then
'                       MsgBox "El Pedido ya fue entregado", vbExclamation, sMensaje
'                       Exit Sub
'                    End If
                    
                     '-----------------------
                    
                    frmPrograma.Show vbModal
                    If wEnter = True Then
                       sPedido = grdGrilla.Columns(0).Text
                       nCambioFecha = sCodigo
                       Isql = "update MPEDIDO set fregistro='" & Format(nCambioFecha, "MM/dd/yyyy HH:mm") & "',  fProgramacion = '" & Format(nCambioFecha, "MM/dd/yyyy HH:mm") & "' where tCodigoPedido='" & sPedido & "'"
                       Cn.Execute Isql
                       RsCabecera.Requery
                    Else
                       nCambioFecha = ""
                    End If
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
        Case Is = 3 ' confirmar entrega
                If RsCabecera.RecordCount > 0 Then
                    sPedido = grdGrilla.Columns(0).Text
                       If grdGrilla.Columns(6).Text = "ENTREGADO" Then
                          MsgBox "El Pedido ya se encuentra Entregado", vbInformation, sMensaje
                          Exit Sub
                       End If
'                       If grdGrilla.Columns(6).Text = "PENDIENTE" And grdGrilla.Columns(7).Text = "NO FACTURADO" Then
'                          MsgBox "El Pedido no tiene documento Asociado", vbInformation, sMensaje
'                          Exit Sub
'                       End If
                        If grdGrilla.Columns(6).Text = "PENDIENTE" And grdGrilla.Columns(7).Text = "NO PAGADO" Then
                          If MsgBox("El Pedido no ha sido Cancelado ¿Desea aún así Entregarlo? ", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                If Supervisor("22") = False Then
                                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                                        Exit Sub
                                End If
                          Else
                                Exit Sub
                          End If
                       End If
                        If grdGrilla.Columns(6).Text = "PENDIENTE" And grdGrilla.Columns(7).Text = "ANTICIPO" Then
                          If MsgBox("El Pedido no ha sido Cancelado ¿Desea aún así Entregarlo? ", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                If Supervisor("22") = False Then
                                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                                        Exit Sub
                                End If
                          Else
                                Exit Sub
                          End If
                       End If
                    If MsgBox("Seguro de Confirmar la entrega del Pedido " & sPedido & "", vbQuestion + vbYesNo) = vbYes Then
                       
                     
                       Dim EstadoDoc As String
                       
                       Isql = "Select * From vDocumentoAgrupado Where tCodigoPedido = '" & sPedido & "'"
                       Set RsDetalleR = Lib.OpenRecordset(Isql, Cn)
                       
                       For i = 1 To RsDetalleR.RecordCount
                             EstadoDoc = RsDetalleR!Estado
                             
                  
'                             If EstadoDoc = "POR COBRAR" Then
'                                MsgBox "No Existe Pago Total del Pedido", vbExclamation, sMensaje
'                                Exit Sub
'                             End If
                             RsDetalleR.MoveNext
                       Next i
                       Cn.Execute "Update MPEDIDO Set lEntregado ='1', tusuarioentregado='" & sUsuario & "', fregentregado=getdate() Where tCodigoPedido='" & sPedido & "'"
                       RsCabecera.Requery
                       MsgBox "Pedido Entregado", vbInformation, sMensaje
                    Else
                       Exit Sub
                    End If
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
                
        Case Is = 4
                Unload Me
        Case Is = 5
                If RsCabecera.RecordCount > 0 Then
                    If grdGrilla.Columns(6).Text <> "ENTREGADO" Then
                          MsgBox "El Pedido no ha sido Entregado", vbInformation, sMensaje
                          Exit Sub
                     End If
                     If Supervisor("22") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                      sPedido = grdGrilla.Columns(0).Text

                    Cn.Execute "Update MPEDIDO Set lEntregado ='0', tusuarioentregado='" & sUsuario & "', fregentregado=getdate() Where tCodigoPedido='" & sPedido & "'"
                    MsgBox "Operación exitosa", vbInformation, sMensaje
                    RsCabecera.Requery
                End If
        Case Is = 6
                If RsCabecera.RecordCount > 0 Then
                    Screen.MousePointer = vbHourglass
                    
                    nPos = grdGrilla.Bookmark
                    Isql = "usp_Inforest_PedidosCentralPedido '" & Format(dtpFecIni.value, "yyyyMMdd 00:00") & "','" & Format(dtpFecFin.value, "yyyyMMdd 23:59") & "'"
                    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
                    Set grdGrilla.DataSource = RsCabecera
                    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
                    grdGrilla.Bookmark = nPos

                     Screen.MousePointer = vbDefault
                End If
                

    End Select
End Sub

Private Sub cmdProcesa_Click()
    Screen.MousePointer = vbHourglass
    
'    Isql = "Select M.tCodigoPedido,V.Descripcion As CanalVenta, T.tNombre + ' ' + T.tApellido As Cliente, CONVERT(NVARCHAR,M.fProgramacion,103) As FechaEntrega, " & _
'           "CONVERT(NVARCHAR,M.fProgramacion,108) As HoraEntrega, SUM(D.nVenta) As Monto,(SELECT EstadoPedido = CASE M.lEntregado WHEN 1 THEN 'ENTREGADO' ELSE 'PENDIENTE' END) As EstadoPedido, ISNULL(T1.Estado,'PAGADO') As Cancelacion, ISNULL(T1.nVenta,0) As SaldoPendiente " & _
'           "From MPEDIDO M INNER JOIN TDELIVERY T ON M.tClienteDelivery = T.tCodigoDelivery INNER JOIN vTipopedido V ON M.tTipoPedido = V.Codigo INNER JOIN DPEDIDO D ON M.tCodigoPedido = D.tCodigoPedido LEFT OUTER JOIN (Select * From vDocumentoAgrupado Where  Estado='POR COBRAR') As T1 ON M.tCodigoPedido = T1.tCodigoPedido " & _
'           "WHERE V.lCanalCentralPedidos = 1 And M.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and M.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' GROUP BY  M.tCodigoPedido, T.tNombre, T.tApellido,M.fProgramacion,V.Descripcion,M.lEntregado,T1.Estado,T1.nVenta ORDER BY FechaEntrega,HoraEntrega"
    Isql = "usp_Inforest_PedidosCentralPedido '" & Format(dtpFecIni.value, "yyyyMMdd 00:00") & "','" & Format(dtpFecFin.value, "yyyyMMdd 23:59") & "'"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    Else
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdGrilla_FilterChange()
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ' Ingrese el Titulo
    Me.Caption = " Correlativo de Pedidos Centro Producción "
    grdGrilla.Caption = Me.Caption
    
    Centrar Me
    nColumna = 0
    
    dtpFecIni.value = FechaServidor()
    dtpFecFin.value = FechaServidor() + 1
    
     If lAccesoDespachoPedido = False Then
        cmdOpcion(2).Enabled = False
        cmdOpcion(3).Enabled = False
        cmdOpcion(5).Enabled = False
    End If

'    Isql = "Select M.tCodigoPedido,V.Descripcion As CanalVenta, T.tNombre + ' ' + T.tApellido As Cliente, CONVERT(NVARCHAR,M.fProgramacion,103) As FechaEntrega, " & _
'           "CONVERT(NVARCHAR,M.fProgramacion,108) As HoraEntrega, SUM(D.nVenta) As Monto,(SELECT EstadoPedido = CASE M.lEntregado WHEN 1 THEN 'ENTREGADO' ELSE 'PENDIENTE' END) As EstadoPedido, ISNULL(T1.Estado,'PAGADO') As Cancelacion, ISNULL(T1.nVenta,0) As SaldoPendiente " & _
'           "From MPEDIDO M INNER JOIN TDELIVERY T ON M.tClienteDelivery = T.tCodigoDelivery INNER JOIN vTipopedido V ON M.tTipoPedido = V.Codigo INNER JOIN DPEDIDO D ON M.tCodigoPedido = D.tCodigoPedido LEFT OUTER JOIN (Select * From vDocumentoAgrupado Where  Estado='POR COBRAR') As T1 ON M.tCodigoPedido = T1.tCodigoPedido " & _
'           "WHERE V.lCanalCentralPedidos = 1 And M.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and M.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' GROUP BY  M.tCodigoPedido, T.tNombre, T.tApellido,M.fProgramacion,V.Descripcion,M.lEntregado,T1.Estado,T1.nVenta ORDER BY FechaEntrega,HoraEntrega"
        Isql = "usp_Inforest_PedidosCentralPedido '" & Format(dtpFecIni.value, "yyyyMMdd 00:00") & "','" & Format(dtpFecFin.value, "yyyyMMdd 23:59") & "'"

    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(9, grdGrilla, "Pedido", 2, "tCodigoPedido", 1200, 2, 0, "", _
                                  "CanalVenta", 2, "CanalVenta", 1500, 2, 0, "", _
                                  "Cliente", 2, "Cliente", 2400, 2, 0, "", _
                                  "FechaEntrega", 2, "FechaEntrega", 1600, 2, 0, "", _
                                  "HoraEntrega", 2, "HoraEntrega", 1600, 2, 0, "", _
                                  "Monto", 2, "Monto", 950, 1, 0, "###,##0.00", _
                                  "EstadoPedido", 2, "EstadoPedido", 1600, 2, 0, "", _
                                  "Cancelacion", 2, "Cancelacion", 1500, 2, 0, "", _
                                  "SaldoPendiente", 2, "SaldoPendiente", 1800, 1, 0, "###,##0.00")
    
        
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub
