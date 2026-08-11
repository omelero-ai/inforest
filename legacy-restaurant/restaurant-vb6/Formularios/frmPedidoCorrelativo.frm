VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPedidoCorrelativo 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7815
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   13695
   Icon            =   "frmPedidoCorrelativo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7788.684
   ScaleMode       =   0  'User
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   13635
      TabIndex        =   12
      Top             =   6435
      Width           =   13695
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
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
         Left            =   30
         TabIndex        =   22
         Top             =   645
         Width           =   5235
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   3195
            TabIndex        =   23
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   51118081
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   720
            TabIndex        =   24
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   51118081
            CurrentDate     =   37539
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
            Left            =   2700
            TabIndex        =   26
            Top             =   300
            Width           =   225
            WordWrap        =   -1  'True
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
            Left            =   90
            TabIndex        =   25
            Top             =   300
            Width           =   405
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   14
         Top             =   45
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmPedidoCorrelativo.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmPedidoCorrelativo.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmPedidoCorrelativo.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmPedidoCorrelativo.frx":1408
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
            Picture         =   "frmPedidoCorrelativo.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmPedidoCorrelativo.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   21
            Top             =   150
            Width           =   2505
         End
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   9
         Left            =   8715
         Picture         =   "frmPedidoCorrelativo.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Reimprime los pedidos a las areas"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   9825
         Picture         =   "frmPedidoCorrelativo.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Anula los Cargos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   10
         Left            =   9270
         Picture         =   "frmPedidoCorrelativo.frx":2A02
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Registra las Propinas"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   10380
         Picture         =   "frmPedidoCorrelativo.frx":2D0C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Despega los Pedidos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   7050
         Picture         =   "frmPedidoCorrelativo.frx":2E0E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Documentos asociados al Pedido"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   5940
         Picture         =   "frmPedidoCorrelativo.frx":2F10
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Detalle del Pedido"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   6495
         Picture         =   "frmPedidoCorrelativo.frx":3012
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Reimprime la Pre Cuenta"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   8160
         Picture         =   "frmPedidoCorrelativo.frx":3544
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar HTML"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   5265
         Picture         =   "frmPedidoCorrelativo.frx":3A76
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Filtrar"
         Top             =   735
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   10935
         Picture         =   "frmPedidoCorrelativo.frx":3B78
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdEmite 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7605
         Picture         =   "frmPedidoCorrelativo.frx":3C6A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Emite"
         Top             =   45
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   6405
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13695
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7140
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12594
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
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
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         Left            =   135
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox ImageBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   11520
      ScaleHeight     =   1185
      ScaleWidth      =   3825
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   3852
   End
End
Attribute VB_Name = "frmPedidoCorrelativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim sPreg     As String
Dim nColumna As Integer
Public nPropina As Double
Public tPropina As String
Dim Correlativo As Integer
Dim tItem As Integer
Dim CorrelaProp As Integer
Dim sMonedaBase As String
Dim i As Integer
Dim rsReporte As Recordset
Dim nPos As Integer
Dim RsImpresion As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean
Dim xTiporecepcion As Integer

Private Sub CmdEmite_Click()
    With grdGrilla.PrintInfo
        'Cabecera
        .SettingsOrientation = 2
        .PageHeaderFont.Italic = True
        .RepeatColumnHeaders = True
        .PreviewMaximize = True
        .SettingsMarginLeft = 400
        .SettingsMarginRight = 400
        .SettingsMarginTop = 400
        .SettingsMarginBottom = 400
        .VariableRowHeight = True
        .PageFooter = "\tPage: \p"
        .PreviewCaption = "Emite de " & Me.Caption
        .PrintPreview
    End With
End Sub

Private Sub cmdExporta_Click()
    cmdSave.Filter = "Hyper Text Markup Language|*.html"
    cmdSave.ShowSave
    If cmdSave.FileName = "" Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    grdGrilla.ExportToFile cmdSave.FileName, False, False
    Screen.MousePointer = vbNormal
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
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & Err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
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

    'Diego L. 05/09/2017
    Dim nPropinaCalc As Double
    Select Case Index
           Case Is = 0  'Detalle
                If RsCabecera.RecordCount > 0 Then
                   sPedido = grdGrilla.Columns(2).Text
                   frmPedidoAnterior.Show vbModal
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
           
           Case Is = 1  'Detalle
                If RsCabecera.RecordCount > 0 Then
                    If Len(grdGrilla.Columns(2).Text) = 10 Then
                        sPedido = grdGrilla.Columns(2).Text
                        frmPedido.cmdOpcion(5).Enabled = False
                        frmPedido.cmdOpcion(6).Enabled = False
                        frmPedido.cmdOpcion(7).Enabled = False
                        frmPedido.Show
                   Else
                        MsgBox "Debe Seleccionar el Pedido a consultar", vbExclamation, sMensaje
                   End If
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
           
           Case Is = 5  'Anular Cargo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
                
                If RsCabecera!tCaja <> sCaja Then
                   MsgBox "No se puede anular el cargo porque pertenece a otra Caja (Caja " & sCaja & ") ", vbCritical, sMensaje
                   Exit Sub
                End If
                
                If RsCabecera!tEstadoPedido <> "05" Then
                   MsgBox "No es un Pedido Cargado", vbExclamation, sMensaje
                   Exit Sub
                End If
                   
                If IIf(IsNull(RsCabecera!Documento), "", RsCabecera!Documento) <> "" Then
                   MsgBox "El Pedido tiene items facturados", vbExclamation, sMensaje
                   Exit Sub
                End If
                   
                If MsgBox("Seguro de Anular el Cargo? ", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                   
                If RsCabecera!tTurno = sTurno Then
                   'Password
                   If Supervisor("05") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                   'Password
                   If MsgBox("El Cargo es de un turno Anterior, deseas continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                      Exit Sub
                   End If

                   If Supervisor("06") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                
                frmKeyBoard.Caption = "Motivo de la Anulación"
                frmKeyBoard.Show vbModal
                If wEnter = False Then
                   Exit Sub
                End If
                 
                If Len(Trim(sDescrip)) = 0 Then
                   MsgBox "Debe ingresar un Motivo", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                On Error GoTo FALLA
                CnInfhotel.BeginTrans
                CnInfhotel.Execute "sp_EliminarPedidoInfhotel '" & RsCabecera!tComanda & "','" & RsCabecera!tPuntoVenta & "'"
                
                
                
'                CnInfhotel.Execute "delete from TCUENTARESERVA where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "delete from TCUENTAHABITACION where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "update MPROPINA set testado='04' where tcomanda='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "delete from DCOMANDA where tComanda='" & RsCabecera!tComanda & "' and tCodigoItem='100000'"
'                CnInfhotel.Execute "update MCOMANDA set TESTADO='04' where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "delete from DCOMANDA where tcomanda='" & RsCabecera!tComanda & "' and tcodigoitem='100000' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "delete from WMCOMANDA where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
'                CnInfhotel.Execute "delete from WDCOMANDA where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & RsCabecera!tPuntoVenta & "'"
                CnInfhotel.CommitTrans
                
                'If CnInfhotel.CommitTrans Then
                    Cn.Execute "delete from MPROPINA where tcodigopedido='" & RsCabecera!tCodigoPedido & "'"
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01' where tCodigoPedido ='" & RsCabecera!tCodigoPedido & "'"
'                Else
'                    CnInfhotel.RollbackTrans
'                    MsgBox ("No se completo la Transaccion!!!!"), vbInformation
'                    Exit Sub
'                End If
                
                nPos = RsCabecera.AbsolutePosition
                RsCabecera.Requery
                RsCabecera.AbsolutePosition = nPos
                Screen.MousePointer = vbDefault
                MsgBox ("Se ha completado la Transaccion!!!!"), vbInformation
                Exit Sub
FALLA:
            CnInfhotel.RollbackTrans
            Screen.MousePointer = vbDefault
            MsgBox ("ERROR: " & Error & " / Volver a Anular cargo"), vbCritical
           Case Is = 6  'Salir
                Unload Me
                      
           Case Is = 8  'Emite / Precuenta
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoPedido = "03" Then
                   MsgBox "Es un pedido Anulado", vbExclamation, sMensaje
                   Exit Sub
                End If
           
'                If Calcular("select case when lcierre=1 then '1' else '0' end  codigo from MTURNO where tTurno='" & RsCabecera!tTurno & "'", Cn) = "0" Then
'                    If Supervisor("14") = False Then
'                      MsgBox "Clave no permitida", vbExclamation, sMensaje
'                      Exit Sub
'                    End If
'
'                Else
'                    MsgBox "El Pedido pertenece a un Turno ya Cerrado", vbInformation, sMensaje
'                    Exit Sub
'                End If

                If Supervisor("12") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
                End If
               

               If lPrecuentaImpresora Then
                  frmPrecuentaImpresora.Show vbModal
                  If Not wEnter Then
                     Exit Sub
                  End If
               Else
                  sCodigo = sPreCuenta
               End If
               sPedido = RsCabecera!tCodigoPedido
               
               Dim nLP As Integer
               nLP = Calcular("select nPrecuenta as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
               If nLimitePrecuenta <> 0 And nLimitePrecuenta <= nLP Then
                  If Supervisor("14") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                   End If
               End If
            
               Screen.MousePointer = vbHourglass
               Dim lPrecuenta As Boolean
               sPropina = IIf(IsNull(RsCabecera!nPropina), 0, RsCabecera!nPropina)
               sMonPropina = IIf(IsNull(RsCabecera!tMonedaPropina), "", RsCabecera!tMonedaPropina)
               lPrecuenta = False
               wEnter = False
               Screen.MousePointer = vbHourglass
               
               'Propiedades
               Dim xSql As String
               If lAlmacen Then
                  Dim RsOp As Recordset
                  Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
                  If RsOp.RecordCount > 0 Then
                     'CASO CORP.FERNANDEZ
                     If BDLink = "1" Then
                        xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                               "nInsumo, nGasto, nManoObra " & _
                               "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                               "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                               "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                               "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                               "Where lNoDescargo = 1"
                    Else
                        xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                            "nInsumo, nGasto, nManoObra " & _
                            "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                            "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                            "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                            "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                            "Where lNoDescargo = 1"
                        End If
                        
                  Else
                     xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                            "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                            "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                            "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
                  End If
               Else
                     xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                            "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                            "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                            "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
               End If
               'Producto Propiedad
               Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador " & _
                      "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
                      "where tCodigoPedido='" & sPedido & "'"
               Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
                 
               'Combo Propiedad
               Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador " & _
                      "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
                      "where tCodigoPedido='" & sPedido & "'"
               Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
               
               Cn.Execute "update MPEDIDO set nPrecuenta = isnull(nPrecuenta,0) + 1 where tCodigoPedido='" & sPedido & "'"
               
'               If RsCabecera!tTipoPedido = "02" Then
'                  Isql = "SELECT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, dbo.DPEDIDO.tItem, TPRODUCTO_2.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nImpuesto1, dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, TPRODUCTO_1.tDetallado AS Combo, dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.nDescuento AS xDescuento, " & _
'                         "dbo.DPEDIDO.nDescuento, dbo.DPEDIDO.nRecargo, dbo.CPEDIDO.nCantidad AS nCantidadCombo, dbo.vDelivery.tApellido, dbo.vDelivery.tNombre, dbo.vDelivery.tDireccion, dbo.vDelivery.tTelefono, dbo.vDelivery.tReferencia, dbo.vDelivery.Zona, dbo.vDelivery.lPuntos, dbo.vDelivery.nDisponible, dbo.vDelivery.tresumido,dbo.MPEDIDO.tPasajero AS tEntrega, dbo.DPEDIDO.tObservacion AS ObservacionDetalle, dbo.vDelivery.tObservacion AS ObservacionCliente " & _
'                         "FROM dbo.vTipoPedido RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vDelivery ON dbo.MPEDIDO.tClienteDelivery = dbo.vDelivery.Codigo ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN " & _
'                         "dbo.CPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_2 ON dbo.DPEDIDO.tCodigoProducto = TPRODUCTO_2.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                         "WHERE dbo.MPEDIDO.tCodigoPedido='" & sPedido & "'"
'               Else
'                  If lPrecuenta Then
'                     Isql = "SELECT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nImpuesto1, dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tObservacion, " & _
'                            "dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, dbo.vTipoPedido.Descripcion AS TipoPedido, CONVERT(nvarchar, dbo.DPEDIDO.fRegistro, 103) AS Fecha, dbo.DPEDIDO.tUsuarioD " & _
'                            "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto LEFT OUTER JOIN dbo.vTipoPedido RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN " & _
'                            "dbo.vSalon RIGHT OUTER JOIN dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
'                            "WHERE dbo.MPEDIDO.tCodigoPedido='" & sPedido & "' ORDER BY dbo.DPEDIDO.fRegistro, dbo.DPEDIDO.tUsuarioD "
'                  Else
'                     If lPrecuentaAgrupada Then
'                         Isql = "SELECT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.nAdulto, TPRODUCTO_2.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.vDelivery.lpuntos, dbo.vDelivery.nDisponible, MAX(dbo.DPEDIDO.tItem) AS tItem, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(dbo.DPEDIDO.nImpuesto1) AS nImpuesto1, SUM(dbo.DPEDIDO.nImpuesto2) AS nImpuesto2, SUM(dbo.DPEDIDO.nImpuesto3) AS nImpuesto3, SUM(dbo.DPEDIDO.nVenta) AS nVenta, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, " & _
'                                "dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.nDescuento AS xDescuento, dbo.vDelivery.Cliente, AVG(dbo.DPEDIDO.nDescuento) AS nDescuento, SUM(dbo.DPEDIDO.nRecargo) AS nRecargo, '' AS Combo " & _
'                                "FROM dbo.vTipoPedido RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vDelivery ON dbo.MPEDIDO.tClienteDelivery = dbo.vDelivery.Codigo ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO AS TPRODUCTO_2 ON dbo.DPEDIDO.tCodigoProducto = TPRODUCTO_2.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                                "WHERE  dbo.MPEDIDO.tCodigoPedido='" & sPedido & "'" & _
'                                "GROUP BY dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.nAdulto, TPRODUCTO_2.tResumido, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, dbo.vTipoPedido.Descripcion, dbo.vMozo.Descripcion, dbo.MPEDIDO.nDescuento, dbo.DPEDIDO.nPrecioOficial , dbo.vDelivery.nDisponible, dbo.vDelivery.lpuntos, dbo.MPEDIDO.tTipoPedido, dbo.vDelivery.Cliente "
'                     Else
'                        Isql = "SELECT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, dbo.DPEDIDO.tItem, TPRODUCTO_2.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nImpuesto1, dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, " & _
'                               "dbo.vTipoPedido.Descripcion AS TipoPedido, TPRODUCTO_1.tDetallado AS Combo, dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.nDescuento AS xDescuento, dbo.DPEDIDO.nDescuento, dbo.DPEDIDO.nRecargo, dbo.MPEDIDO.tTipoPedido, dbo.vDelivery.lPuntos, dbo.vDelivery.nDisponible, dbo.MPEDIDO.nAdulto, dbo.CPEDIDO.nCantidad AS nCantidadCombo " & _
'                               "FROM dbo.vTipoPedido RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vDelivery ON dbo.MPEDIDO.tClienteDelivery = dbo.vDelivery.Codigo ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.vSalon RIGHT OUTER JOIN dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.CPEDIDO LEFT OUTER JOIN " & _
'                               "dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_2 ON dbo.DPEDIDO.tCodigoProducto = TPRODUCTO_2.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                               "WHERE dbo.MPEDIDO.tCodigoPedido='" & sPedido & "' order by dbo.DPEDIDO.tItem"
'                     End If
'                  End If
'               End If
               
                 xTiporecepcion = Calcular("Select isnull(tTiporecepcion,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                
                If RsCabecera!tTipoPedido = "02" Or (WEBAP = True And xTiporecepcion <> 0) Then
'                   Isql = "select * from vPreCuentaDelivery WHERE Codigo='" & sPedido & "' ORDER BY tItem, tItemCombo"
                    Isql = "select t1.*, t2.fProgramacion from vPreCuentaDelivery T1 inner join MPEDIDO t2 on t1.Codigo = t2.tCodigoPedido  WHERE T1.Codigo='" & sPedido & "' ORDER BY T1.tItem, T1.tItemCombo"
                Else
                   If lPrecuenta Then 'Precuenta con mozos
                      Isql = "select * from vPrecuentaDetallada WHERE Codigo='" & sPedido & "'"
                   Else
                      If lPrecuentaAgrupada Then
                         Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                      Else
                         Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                      End If
                   End If
                End If
            
                If lBuscarPedidoBarraMesero Then
                    BarraCodigo sPedido, ImageBarra
                    'Exit Sub
                End If
                         
               Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
               If RsImpresion.RecordCount = 0 Then
                  LimpiaRs
                  MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
               Else
                  If RsCabecera!tTipoPedido = "02" Then
                     ImprimeDelivery RsImpresion, RsProductoPropiedad, RsComboPropiedad, Nothing, Nothing, ImageBarra
                  Else
                     If lPrecuenta Then
                        ImprimePreCuentaDetallada RsImpresion, sCodigo, Nothing, Nothing, ImageBarra
                     Else
                        If lInfhotel Then
                           sTipoComanda = IIf(IsNull(RsCabecera!tTipoComanda), "", RsCabecera!tTipoComanda)
                           sReserva = IIf(IsNull(RsCabecera!tReserva), "", RsCabecera!tReserva)
                           sFichaPasajero = IIf(IsNull(RsCabecera!tFichaPasajero), "", RsCabecera!tFichaPasajero)
                           ImprimeInfhotel RsImpresion, sCodigo
                        Else
                            If lPrecuentaNoValorizada Then
                                ImprimePrecuentaNoValorizada RsImpresion, sCodigo, Nothing, Nothing, ImageBarra
                            Else
                                ImprimePreCuenta RsImpresion, sCodigo, Nothing, Nothing, ImageBarra
                            End If
                        End If
                     End If
                  End If
                  Cn.Execute "update MPEDIDO set fRegCuenta = getDate() where tCodigoPedido='" & sPedido & "'"
                  
               End If
               LimpiaRs
               
                
           Case Is = 9  'Impresion de Pedidos
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoPedido = "03" Then
                   MsgBox "Es un pedido Anulado", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If Calcular("select case when lcierre=1 then '1' else '0' end  codigo from MTURNO where tTurno='" & RsCabecera!tTurno & "'", Cn) = "0" Then
                    If Supervisor("14") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                    End If
                Else
                    MsgBox "El Pedido pertenece a un Turno ya Cerrado", vbInformation, sMensaje
                    Exit Sub
                End If
               
                
                Screen.MousePointer = vbHourglass
                Dim RsArea As Recordset
                Dim i As Integer
                
                Dim xLP As Integer
                
                If RsCabecera.RecordCount > 0 Then
                   sPedido = grdGrilla.Columns(2).Text

                End If
                
                xLP = Calcular("select nReimpresion as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
                'If nLimitePrecuenta <> 0 And nLimitePrecuenta <= xLP Then
                If nLimiteReimpresion <> 0 And nLimiteReimpresion <= xLP Then
                   If Supervisor("14") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                
                Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                If lAlmacen Then
                   Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
                   
                   If RsOp.RecordCount > 0 Then
                      'CASO CORP.FERNANDEZ
                    If BDLink = "1" Then
                        xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                               "nInsumo, nGasto, nManoObra " & _
                               "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                               "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                               "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                               "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                               "Where lNoDescargo = 1"
                        Else
                            xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                                 "nInsumo, nGasto, nManoObra " & _
                                 "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                                 "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                                 "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                                 "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto " & _
                                 "Where lNoDescargo = 1"
                        End If
                   Else
                      xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                             "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                             "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                             "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
                   End If
                Else
                      xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                             "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                             "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                             "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
                End If
   
                'Producto Propiedad
                Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador, isnull(tproductopropiedad.ncantidad,1) ncantidad   " & _
                       "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
                       "where tCodigoPedido = '" & RsCabecera!tCodigoPedido & "'"
                Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
                   
                'Combo Propiedad
                Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador " & _
                       "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
                       "where tCodigoPedido = '" & RsCabecera!tCodigoPedido & "'"
                 Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
                                
                Isql = "select * from vPedido " & _
                       "Where Codigo = '" & RsCabecera!tCodigoPedido & "' And lImprimeArea = 1 and (lImprime = 1 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 1 )) " & _
                       "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo"
                       '"ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                       
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                            
                If Not RsImpresion.EOF Then
                   RsArea.MoveFirst
                   For i = 1 To RsArea.RecordCount
                       RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                       If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                          If RsImpresion.RecordCount <> 0 Then
                             RsImpresion.MoveFirst
                             ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, True, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                          End If
                       End If
                       RsArea.MoveNext
                   Next i
                End If
                Cn.Execute "Update DPEDIDO Set nEnvio = isnull(nEnvio,0) + 1 where tCodigoPedido = '" & RsCabecera!tCodigoPedido & "'" ' and lImprimeArea = 1 and tArea='" & sCodigo & "'"
                Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                LimpiaRs
                      
           Case Is = 10  'Cambio de Propina
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
                
                If RsCabecera!tCaja <> sCaja Then
                   MsgBox "No se modificar porque pertenece a otra Caja (Caja " & sCaja & ") ", vbCritical, sMensaje
                   Exit Sub
                End If
                
                If RsCabecera!tTurno <> sTurno Then
                   MsgBox "No se puede modificar la propina de otro turno", vbCritical, sMensaje
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoPedido = "03" Then
                   MsgBox "Es un pedido Anulado", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoPedido <> "05" Then
                   MsgBox "No es un Pedido Cargado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If RsCabecera!nPropina <> "" Then
                   MsgBox "El Pedido ya tiene propina", vbExclamation, sMensaje
                   Exit Sub
                End If
                wEnter = False
                
                nPropina = 0
                frmCambioPropina.Show vbModal
                Screen.MousePointer = vbHourglass
                sMonedaBase = Calcular("select tMoneda as Codigo from TPARAMETRO", CnInfhotel)
                If wEnter Then
                   CorrelaProp = Calcular("select max(ncorrela) as codigo from mpropina", CnInfhotel) + 1
                   
                   Isql = "Update MCOMANDA set ncorrelaprop=" & CorrelaProp & " where tcomanda='" & RsCabecera!tComanda & "'"
                   CnInfhotel.Execute Isql
                   tItem = Calcular("select max(titem) as codigo from dcomanda where tcomanda='" & RsCabecera!tComanda & "'", CnInfhotel) + 1
                   Isql = "Insert Into dcomanda " & _
                          "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                          "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                          "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                          "values('" & RsCabecera!tComanda & "','" & RsCabecera!tPuntoVenta & "','00','" & tItem & "','1','" & RsCabecera!tCodigoPedido & "', " & _
                          "'100000'," & nPropina & ",'1'," & nPropina & ",'','" & RsCabecera!tReserva & "','" & RsCabecera!tHabitacion & "','', " & _
                          " '','','','" & xUsuario & "',getdate(),getdate())"
                   CnInfhotel.Execute Isql
                
                   'Graba en el mpropina de inforest
                   Isql = "insert into MPROPINA " & _
                          "(tcodigopedido, fregistro, tmoneda, nmonto, tusuario, tComanda) " & _
                          "values('" & RsCabecera!tCodigoPedido & "',getdate(),'" & tPropina & "'," & nPropina & ",'" & sUsuario & "', '" & RsCabecera!tComanda & "')"
                   Cn.Execute Isql
                   
                   'Graba en el mpropina de infhotel
                    Isql = "Insert Into MPROPINA " & _
                       "(ncorrela,tcodigoreserva,tnumerohabitacion,tcomanda,tcodigoitem,tmoneda,nmonto,tdocumento,tresponsable,testado,ffecha,tusuario) " & _
                       "values(" & CorrelaProp & ",'" & RsCabecera!tReserva & "','" & RsCabecera!tHabitacion & "','" & RsCabecera!tComanda & "','100000','" & tPropina & "'," & nPropina & ", " & _
                       "'','" & Mid(RsCabecera!tMozo, 2, 3) & "','01',getdate(),'" & xUsuario & "')"
                    CnInfhotel.Execute Isql
                   
                   If RsCabecera!tHabitacion <> "" Then
                      i = Calcular("select max(titem) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & RsCabecera!tReserva & "'and tcomanda='" & RsCabecera!tComanda & "'", CnInfhotel) + 1
                      Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & RsCabecera!tReserva & "'", CnInfhotel) + 1
                      
                    If sMonedaBase = tPropina Then
                         nPropinaCalc = nPropina
                    Else
                     If (sMonedaBase = "02" And tPropina = "01") Then
                            nPropinaCalc = CDbl(nPropina) / nTC
                     Else
                            nPropinaCalc = CDbl(nPropina) * nTC
                     End If
                    End If
                      
                      Isql = "Insert into TCUENTAHABITACION " & _
                         "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                         " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                         " Values('" & RsCabecera!tReserva & "','" & RsCabecera!tHabitacion & "',getdate(),getdate(),'','','" & RsCabecera!tComanda & "','" & RsCabecera!tCodigoPedido & "','100000'," & nPropinaCalc & ",'1','00', " & _
                         " " & nPropinaCalc & "," & Correlativo & ",'" & RsCabecera!tPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                      CnInfhotel.Execute Isql
'                      If sMonedaBase = tPropina Then
'                        Isql = "UPDATE tCuentaHabitacion " & _
'                              " SET nPrecioUnitario=" & CDbl(nPropina) & "," & _
'                              " nCantidad=1," & _
'                              " nTotal=" & CDbl(nPropina) & " " & _
'                              " WHERE ncorrelaprop='" & CorrelaProp & "'"
'                         CnInfhotel.Execute Isql
'                      Else
'                        If (sMonedaBase = "02" And tPropina = "01") Then
'                            'La moneda base esta en $ y la propina esta en S/.
'                            Isql = "UPDATE tcuentahabitacion " & _
'                                 " SET nPrecioUnitario=" & CDbl(nPropina) / nTC & "," & _
'                                 " nCantidad=1," & _
'                                 " ntotal=" & CDbl(nPropina) / nTC & "" & _
'                                 " WHERE nCorrelaProp='" & CorrelaProp & "'"
'                            CnInfhotel.Execute Isql
'                        Else
'                            'La moneda base esta en S/. y la propina esta en $
'                            Isql = "UPDATE tcuentahabitacion " & _
'                                 " SET nPrecioUnitario=" & CDbl(nPropina) * nTC & "," & _
'                                 " nCantidad=1," & _
'                                 " ntotal=" & CDbl(nPropina) * nTC & "" & _
'                                 " WHERE nCorrelaProp='" & CorrelaProp & "'"
'                            CnInfhotel.Execute Isql
'                        End If
'                      End If
                    ElseIf RsCabecera!tReserva <> "" Then
                            i = Calcular("select max(titem) as Codigo from TCUENTARESERVA where tCodigoReserva='" & RsCabecera!tReserva & "'and tcomanda='" & RsCabecera!tComanda & "'", CnInfhotel) + 1
                            Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & RsCabecera!tReserva & "'", CnInfhotel) + 1
                            
                            If sMonedaBase = tPropina Then
                                 nPropinaCalc = nPropina
                            Else
                             If (sMonedaBase = "02" And tPropina = "01") Then
                                    nPropinaCalc = CDbl(nPropina) / nTC
                             Else
                                    nPropinaCalc = CDbl(nPropina) * nTC
                             End If
                            End If
                            
                            Isql = "insert into tcuentareserva " & _
                               "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado,ttipo,tHotel, " & _
                               " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                               " values('" & RsCabecera!tReserva & "','" & RsCabecera!tHabitacion & "',getdate(),getdate(),'" & RsCabecera!tComanda & "','" & RsCabecera!tCodigoPedido & "','100000'," & nPropinaCalc & ",'1','','','00', " & _
                               " " & nPropinaCalc & "," & Correlativo & ",'" & RsCabecera!tPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                            CnInfhotel.Execute Isql
                            
                                                        
'                            If sMonedaBase = tPropina Then
'                                Isql = "UPDATE tCuentaReserva " & _
'                                   " SET nPrecioUnitario=" & CDbl(nPropina) & "," & _
'                                   " nCantidad=1," & _
'                                   " nTotal=" & CDbl(nPropina) & " " & _
'                                   " WHERE ncorrelaprop='" & CorrelaProp & "'"
'                                CnInfhotel.Execute Isql
'                            Else
'                             If (sMonedaBase = "02" And tPropina = "01") Then
'                                'La moneda base esta en $ y la propina esta en S/.
'                                 Isql = "UPDATE tcuentareserva " & _
'                                      " SET nPrecioUnitario=" & CDbl(nPropina) / nTC & "," & _
'                                      " nCantidad=1," & _
'                                      " ntotal=" & CDbl(nPropina) / nTC & "" & _
'                                      " WHERE nCorrelaProp='" & CorrelaProp & "'"
'                                 CnInfhotel.Execute Isql
'                             Else
'                                 'La moneda base esta en S/. y la propina esta en $
'                                 Isql = "UPDATE tcuentaReserva" & _
'                                      " SET nPrecioUnitario=" & CDbl(nPropina) * nTC & "," & _
'                                      " nCantidad=1," & _
'                                      " ntotal=" & CDbl(nPropina) * nTC & "" & _
'                                      " WHERE nCorrelaProp='" & CorrelaProp & "'"
'                                 CnInfhotel.Execute Isql
'                             End If
'                            End If
                       End If
               nPos = RsCabecera.AbsolutePosition
               RsCabecera.Requery
               RsCabecera.AbsolutePosition = nPos
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
                
           
           Case Is = 11  'Despegar un pedido
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If RsCabecera!tEstadoPedido <> "01" Then
                   MsgBox "El pedido no se encuentra pegado", vbExclamation, sMensaje
                   Exit Sub
                End If
                    
                Dim nContador As Integer
                nContador = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where tCodigoPedido='" & RsCabecera!tCodigoPedido & "' and isnull(tFacturado,'')<>'P'", Cn)
                If nContador > 0 Then
                   MsgBox "El pedido no se encuentra pegado, tienes items por facturar y/o cancelar", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If MsgBox("Deseas Despegar el Pedido? ", vbExclamation + vbYesNo, sMensaje) = vbYes Then
                   Cn.Execute "update MPEDIDO set tEstadoPedido='02' where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                   nPos = RsCabecera.AbsolutePosition
                   RsCabecera.Requery
                   RsCabecera.AbsolutePosition = nPos
                End If
    End Select
End Sub

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
   
    Isql = "select *, str(round(CAST(nVenta AS NUMERIC(18,3)),2),10,2)  as xVenta, convert(varchar, fFecha, 103) as xFecha, str(nAdulto,10,2) as xAdulto " & _
           "from vPedidoCorrelativo " & _
           "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' order by tCodigoPedido"
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    If Not lInfhotel Then
       cmdOpcion(5).Enabled = False
       cmdOpcion(10).Enabled = False
    End If
     
    ' Ingrese el Titulo
    Me.Caption = " Correlativo de Pedidos "
    grdGrilla.Caption = Me.Caption
    
    Centrar Me
    nColumna = 0
    dtpFecIni.Value = FechaServidor() - nDias
    dtpFecFin.Value = FechaServidor()

    ' Ingrese la Vista
    Isql = "select *, str(round(CAST(nVenta AS NUMERIC(18,3)),2),10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, str(nAdulto,10,2) as xAdulto " & _
           "from vPedidoCorrelativo " & _
           "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' order by tCodigoPedido"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    If lInfhotel Then
       Call ConfGrilla(18, grdGrilla, "Caja", 2, "tCaja", 500, 2, 0, "", _
                                      "Turno", 2, "tTurno", 1200, 0, 0, "", _
                                      "Pedido", 2, "tCodigoPedido", 1200, 0, 0, "", _
                                      "Comanda", 2, "tComanda", 1100, 0, 0, "", _
                                      "Fecha", 2, "xFecha", 1000, 0, 0, "dd/mm/yyyy", _
                                      "Mesa", 2, "Mesa", 1100, 0, 0, "", _
                                      "Observación", 2, "tObservacion", 1200, 0, 0, "", _
                                      Mesero, 2, "Mozo", 1200, 0, 0, "", _
                                      "Monto", 2, "xVenta", 950, 1, 0, "###,##0.00", _
                                      "Estado", 2, "Estado", 1200, 0, 0, "", _
                                      "Documento", 2, "Documento", 1500, 0, 0, "", _
                                      "Reserva", 2, "tReserva", 1500, 0, 0, "", _
                                      "Habitacion", 2, "tHabitacion", 1500, 0, 0, "", _
                                      "Pasajero", 2, "tPasajero", 1500, 0, 0, "", _
                                      "Usu.Emi", 2, "tUsuario", 1150, 0, 0, "", _
                                      "Motivo Anulación", 2, "tObservacionAnulado", 2600, 0, 0, "", _
                                      "Mon", 2, "tMonedaPropina", 750, 0, 0, "", _
                                      "Propina", 2, "nPropina", 950, 1, 0, "###,##0.00")
    Else
       Call ConfGrilla(14, grdGrilla, "Caja", 2, "tCaja", 500, 2, 0, "", _
                                      "Turno", 2, "tTurno", 1100, 0, 0, "", _
                                      "Pedido", 2, "tCodigoPedido", 1100, 0, 0, "", _
                                      "Fecha", 2, "xFecha", 1100, 0, 0, "dd/mm/yyyy", _
                                      "Tipo Pedido", 2, "TipoPedido", 1100, 0, 0, "", _
                                      "Mesa", 2, "Mesa", 1200, 0, 0, "", _
                                      Mesero, 2, "Mozo", 1300, 0, 0, "", _
                                      "Pax", 2, "xAdulto", 500, 1, 0, "", _
                                      "Observación", 2, "tObservacion", 1400, 0, 0, "", _
                                      "Monto", 2, "xVenta", 1000, 1, 0, "###,##0.00", _
                                      "Estado", 2, "Estado", 1100, 0, 0, "", _
                                      "Documento", 2, "Documento", 1600, 0, 0, "", _
                                      "Usu.Emi", 2, "tUsuario", 1000, 0, 0, "", _
                                      "Motivo Anulación", 2, "tObservacionAnulado", 1500, 0, 0, "")
    End If

    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsCabecera = Nothing
    Set frmPedidoCorrelativo = Nothing
End Sub

Private Sub grdGrilla_DblClick()
    cmdOpcion_Click (1)
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

