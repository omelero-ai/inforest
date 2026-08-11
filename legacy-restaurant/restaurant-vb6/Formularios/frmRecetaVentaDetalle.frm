VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecetaVentaDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receta de Venta"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecetaVentaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   " 00001 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8445
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12255
      Begin VB.Frame fraObservacion 
         Caption         =   " Observación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   90
         TabIndex        =   4
         Top             =   6750
         Width           =   11955
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
            Index           =   7
            Left            =   10500
            Picture         =   "frmRecetaVentaDetalle.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   780
            Width           =   1245
         End
         Begin VB.TextBox txtObservacion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   300
            Width           =   9975
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   3345
         Left            =   90
         TabIndex        =   3
         Top             =   1110
         Width           =   11955
         Begin TrueOleDBGrid80.TDBGrid grdGrilla 
            Height          =   3030
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   210
            Width           =   11715
            _ExtentX        =   20664
            _ExtentY        =   5345
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0).NumberFormat=   "True/False"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerStyle=   2
            Splits(0).DividerColor=   32768
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "Correlativo"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
            _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(9)   =   ":id=4,.fontname=Arial"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Arial"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
            _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(43)  =   "Named:id=34:Heading"
            _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=34,.fontname=Arial"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
            _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(53)  =   ":id=36,.fontname=Arial"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
            _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   210
         Left            =   10980
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   382
         Value           =   1  'Checked
         Width           =   855
      End
      Begin TabDlg.SSTab TabCostos 
         Height          =   2145
         Left            =   120
         TabIndex        =   11
         Top             =   4560
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   3784
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "n/d"
         TabPicture(0)   =   "frmRecetaVentaDetalle.frx":047C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(9)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblInsumo(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(10)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblCosto(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label(11)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblGasto(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblMObra(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label(2)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label(13)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label(14)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label(15)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label(16)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblUtilidad(0)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lblSugeridoS(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "lblVentaS(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "lblNetoS(0)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label(5)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "lblUtilidadReal(0)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label(4)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label(8)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label(17)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label(18)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label(19)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "lblTCambio(0)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "lblSugeridoD(0)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "lblNetoD(0)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "lblVentaD(0)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtGasto(0)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtMObra(0)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtUtilidad(0)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).ControlCount=   32
         TabCaption(1)   =   "n/d"
         TabPicture(1)   =   "frmRecetaVentaDetalle.frx":0498
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblVentaD(1)"
         Tab(1).Control(1)=   "lblNetoD(1)"
         Tab(1).Control(2)=   "lblSugeridoD(1)"
         Tab(1).Control(3)=   "lblTCambio(1)"
         Tab(1).Control(4)=   "Label(6)"
         Tab(1).Control(5)=   "Label(7)"
         Tab(1).Control(6)=   "Label(20)"
         Tab(1).Control(7)=   "Label(21)"
         Tab(1).Control(8)=   "Label(22)"
         Tab(1).Control(9)=   "lblUtilidadReal(1)"
         Tab(1).Control(10)=   "Label(23)"
         Tab(1).Control(11)=   "lblNetoS(1)"
         Tab(1).Control(12)=   "lblVentaS(1)"
         Tab(1).Control(13)=   "lblSugeridoS(1)"
         Tab(1).Control(14)=   "lblUtilidad(1)"
         Tab(1).Control(15)=   "Label(24)"
         Tab(1).Control(16)=   "Label(25)"
         Tab(1).Control(17)=   "Label(26)"
         Tab(1).Control(18)=   "Label(27)"
         Tab(1).Control(19)=   "Label(28)"
         Tab(1).Control(20)=   "Label(29)"
         Tab(1).Control(21)=   "lblMObra(1)"
         Tab(1).Control(22)=   "lblGasto(1)"
         Tab(1).Control(23)=   "Label(30)"
         Tab(1).Control(24)=   "lblCosto(1)"
         Tab(1).Control(25)=   "Label(31)"
         Tab(1).Control(26)=   "Label(32)"
         Tab(1).Control(27)=   "lblInsumo(1)"
         Tab(1).Control(28)=   "Label(33)"
         Tab(1).Control(29)=   "txtUtilidad(1)"
         Tab(1).Control(30)=   "txtMObra(1)"
         Tab(1).Control(31)=   "txtGasto(1)"
         Tab(1).ControlCount=   32
         TabCaption(2)   =   "n/d"
         TabPicture(2)   =   "frmRecetaVentaDetalle.frx":04B4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblVentaD(2)"
         Tab(2).Control(1)=   "lblNetoD(2)"
         Tab(2).Control(2)=   "lblSugeridoD(2)"
         Tab(2).Control(3)=   "lblTCambio(2)"
         Tab(2).Control(4)=   "Label(34)"
         Tab(2).Control(5)=   "Label(35)"
         Tab(2).Control(6)=   "Label(36)"
         Tab(2).Control(7)=   "Label(37)"
         Tab(2).Control(8)=   "Label(38)"
         Tab(2).Control(9)=   "lblUtilidadReal(2)"
         Tab(2).Control(10)=   "Label(59)"
         Tab(2).Control(11)=   "lblNetoS(2)"
         Tab(2).Control(12)=   "lblVentaS(2)"
         Tab(2).Control(13)=   "lblSugeridoS(2)"
         Tab(2).Control(14)=   "lblUtilidad(2)"
         Tab(2).Control(15)=   "Label(60)"
         Tab(2).Control(16)=   "Label(61)"
         Tab(2).Control(17)=   "Label(62)"
         Tab(2).Control(18)=   "Label(63)"
         Tab(2).Control(19)=   "Label(64)"
         Tab(2).Control(20)=   "Label(65)"
         Tab(2).Control(21)=   "lblMObra(2)"
         Tab(2).Control(22)=   "lblGasto(2)"
         Tab(2).Control(23)=   "Label(66)"
         Tab(2).Control(24)=   "lblCosto(2)"
         Tab(2).Control(25)=   "Label(67)"
         Tab(2).Control(26)=   "Label(68)"
         Tab(2).Control(27)=   "lblInsumo(2)"
         Tab(2).Control(28)=   "Label(69)"
         Tab(2).Control(29)=   "txtUtilidad(2)"
         Tab(2).Control(30)=   "txtMObra(2)"
         Tab(2).Control(31)=   "txtGasto(2)"
         Tab(2).ControlCount=   32
         TabCaption(3)   =   "n/d"
         TabPicture(3)   =   "frmRecetaVentaDetalle.frx":04D0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lblVentaD(3)"
         Tab(3).Control(1)=   "lblNetoD(3)"
         Tab(3).Control(2)=   "lblSugeridoD(3)"
         Tab(3).Control(3)=   "lblTCambio(3)"
         Tab(3).Control(4)=   "Label(70)"
         Tab(3).Control(5)=   "Label(71)"
         Tab(3).Control(6)=   "Label(72)"
         Tab(3).Control(7)=   "Label(73)"
         Tab(3).Control(8)=   "Label(74)"
         Tab(3).Control(9)=   "lblUtilidadReal(3)"
         Tab(3).Control(10)=   "Label(75)"
         Tab(3).Control(11)=   "lblNetoS(3)"
         Tab(3).Control(12)=   "lblVentaS(3)"
         Tab(3).Control(13)=   "lblSugeridoS(3)"
         Tab(3).Control(14)=   "lblUtilidad(3)"
         Tab(3).Control(15)=   "Label(76)"
         Tab(3).Control(16)=   "Label(77)"
         Tab(3).Control(17)=   "Label(78)"
         Tab(3).Control(18)=   "Label(79)"
         Tab(3).Control(19)=   "Label(80)"
         Tab(3).Control(20)=   "Label(81)"
         Tab(3).Control(21)=   "lblMObra(3)"
         Tab(3).Control(22)=   "lblGasto(3)"
         Tab(3).Control(23)=   "Label(82)"
         Tab(3).Control(24)=   "lblCosto(3)"
         Tab(3).Control(25)=   "Label(83)"
         Tab(3).Control(26)=   "Label(84)"
         Tab(3).Control(27)=   "lblInsumo(3)"
         Tab(3).Control(28)=   "Label(85)"
         Tab(3).Control(29)=   "txtUtilidad(3)"
         Tab(3).Control(30)=   "txtMObra(3)"
         Tab(3).Control(31)=   "txtGasto(3)"
         Tab(3).ControlCount=   32
         TabCaption(4)   =   "n/d"
         TabPicture(4)   =   "frmRecetaVentaDetalle.frx":04EC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblVentaD(4)"
         Tab(4).Control(1)=   "lblNetoD(4)"
         Tab(4).Control(2)=   "lblSugeridoD(4)"
         Tab(4).Control(3)=   "lblTCambio(4)"
         Tab(4).Control(4)=   "Label(86)"
         Tab(4).Control(5)=   "Label(87)"
         Tab(4).Control(6)=   "Label(88)"
         Tab(4).Control(7)=   "Label(89)"
         Tab(4).Control(8)=   "Label(90)"
         Tab(4).Control(9)=   "lblUtilidadReal(4)"
         Tab(4).Control(10)=   "Label(91)"
         Tab(4).Control(11)=   "lblNetoS(4)"
         Tab(4).Control(12)=   "lblVentaS(4)"
         Tab(4).Control(13)=   "lblSugeridoS(4)"
         Tab(4).Control(14)=   "lblUtilidad(4)"
         Tab(4).Control(15)=   "Label(92)"
         Tab(4).Control(16)=   "Label(93)"
         Tab(4).Control(17)=   "Label(94)"
         Tab(4).Control(18)=   "Label(95)"
         Tab(4).Control(19)=   "Label(96)"
         Tab(4).Control(20)=   "Label(97)"
         Tab(4).Control(21)=   "lblMObra(4)"
         Tab(4).Control(22)=   "lblGasto(4)"
         Tab(4).Control(23)=   "Label(98)"
         Tab(4).Control(24)=   "lblCosto(4)"
         Tab(4).Control(25)=   "Label(99)"
         Tab(4).Control(26)=   "Label(100)"
         Tab(4).Control(27)=   "lblInsumo(4)"
         Tab(4).Control(28)=   "Label(101)"
         Tab(4).Control(29)=   "txtUtilidad(4)"
         Tab(4).Control(30)=   "txtMObra(4)"
         Tab(4).Control(31)=   "txtGasto(4)"
         Tab(4).ControlCount=   32
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   4
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   186
            Top             =   840
            Width           =   765
         End
         Begin VB.TextBox txtMObra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   4
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   185
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Index           =   4
            Left            =   -67200
            MaxLength       =   7
            TabIndex        =   184
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   3
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   154
            Top             =   840
            Width           =   765
         End
         Begin VB.TextBox txtMObra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   3
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   153
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Index           =   3
            Left            =   -67200
            MaxLength       =   7
            TabIndex        =   152
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   2
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   122
            Top             =   840
            Width           =   765
         End
         Begin VB.TextBox txtMObra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   2
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   121
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Index           =   2
            Left            =   -67200
            MaxLength       =   7
            TabIndex        =   120
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   1
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   90
            Top             =   840
            Width           =   765
         End
         Begin VB.TextBox txtMObra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   1
            Left            =   -71760
            MaxLength       =   6
            TabIndex        =   89
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Index           =   1
            Left            =   -67200
            MaxLength       =   7
            TabIndex        =   88
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txtUtilidad 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Index           =   0
            Left            =   7800
            MaxLength       =   7
            TabIndex        =   76
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txtMObra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   0
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   65
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtGasto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   0
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   64
            Top             =   840
            Width           =   765
         End
         Begin VB.TextBox txtPrecioCosto5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtInsumo5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtGasto5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtManoObra5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtInsumo4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtGasto4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtManoObra4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtInsumo3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtGasto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtManoObra3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtPrecioCosto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtInsumo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   " "
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtGasto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox txtManoObra2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1530
            Width           =   975
         End
         Begin VB.TextBox txtUtilidad2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            Height          =   285
            Left            =   -73620
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   975
         End
         Begin MSChart20Lib.MSChart mscGrafico2 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmRecetaVentaDetalle.frx":0508
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico3 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmRecetaVentaDetalle.frx":2DA9
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico4 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmRecetaVentaDetalle.frx":562E
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin MSChart20Lib.MSChart mscGrafico5 
            Height          =   2550
            Left            =   -72585
            OleObjectBlob   =   "frmRecetaVentaDetalle.frx":7EB3
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   360
            Width           =   3195
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo de Insumos  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   101
            Left            =   -74880
            TabIndex        =   215
            Top             =   540
            Width           =   1770
         End
         Begin VB.Label lblInsumo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -73080
            TabIndex        =   214
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "&Gastos Generales  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   100
            Left            =   -74880
            TabIndex        =   213
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Costo M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   99
            Left            =   -74880
            TabIndex        =   212
            Top             =   1620
            Width           =   1770
         End
         Begin VB.Label lblCosto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Index           =   4
            Left            =   -73080
            TabIndex        =   211
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mano de &Obra M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   98
            Left            =   -74880
            TabIndex        =   210
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblGasto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -73080
            TabIndex        =   209
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblMObra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -73080
            TabIndex        =   208
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   97
            Left            =   -70920
            TabIndex        =   207
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   96
            Left            =   -70920
            TabIndex        =   206
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Utilidad Sugerida M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   95
            Left            =   -70320
            TabIndex        =   205
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Sugerido M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   94
            Left            =   -70560
            TabIndex        =   204
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta Neto M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   93
            Left            =   -70320
            TabIndex        =   203
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   92
            Left            =   -69960
            TabIndex        =   202
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label lblUtilidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -68640
            TabIndex        =   201
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -68640
            TabIndex        =   200
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblVentaS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   4
            Left            =   -68640
            TabIndex        =   199
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblNetoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   4
            Left            =   -68640
            TabIndex        =   198
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   91
            Left            =   -66360
            TabIndex        =   197
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblUtilidadReal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -67200
            TabIndex        =   196
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   90
            Left            =   -66360
            TabIndex        =   195
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   89
            Left            =   -65880
            TabIndex        =   194
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   88
            Left            =   -65115
            TabIndex        =   193
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   87
            Left            =   -65115
            TabIndex        =   192
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   86
            Left            =   -65115
            TabIndex        =   191
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label lblTCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -64560
            TabIndex        =   190
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   -64560
            TabIndex        =   189
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblNetoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   4
            Left            =   -64560
            TabIndex        =   188
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblVentaD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   4
            Left            =   -64560
            TabIndex        =   187
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo de Insumos  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   85
            Left            =   -74880
            TabIndex        =   183
            Top             =   540
            Width           =   1770
         End
         Begin VB.Label lblInsumo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -73080
            TabIndex        =   182
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "&Gastos Generales  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   84
            Left            =   -74880
            TabIndex        =   181
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Costo M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   -74880
            TabIndex        =   180
            Top             =   1620
            Width           =   1770
         End
         Begin VB.Label lblCosto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Index           =   3
            Left            =   -73080
            TabIndex        =   179
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mano de &Obra M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   82
            Left            =   -74880
            TabIndex        =   178
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblGasto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -73080
            TabIndex        =   177
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblMObra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -73080
            TabIndex        =   176
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   81
            Left            =   -70920
            TabIndex        =   175
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   80
            Left            =   -70920
            TabIndex        =   174
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Utilidad Sugerida M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   79
            Left            =   -70320
            TabIndex        =   173
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Sugerido M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   78
            Left            =   -70560
            TabIndex        =   172
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta Neto M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   77
            Left            =   -70320
            TabIndex        =   171
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   76
            Left            =   -69960
            TabIndex        =   170
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label lblUtilidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -68640
            TabIndex        =   169
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -68640
            TabIndex        =   168
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblVentaS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   3
            Left            =   -68640
            TabIndex        =   167
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblNetoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   3
            Left            =   -68640
            TabIndex        =   166
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   75
            Left            =   -66360
            TabIndex        =   165
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblUtilidadReal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -67200
            TabIndex        =   164
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   74
            Left            =   -66360
            TabIndex        =   163
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   73
            Left            =   -65880
            TabIndex        =   162
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   72
            Left            =   -65115
            TabIndex        =   161
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   71
            Left            =   -65115
            TabIndex        =   160
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   70
            Left            =   -65115
            TabIndex        =   159
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label lblTCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -64560
            TabIndex        =   158
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   -64560
            TabIndex        =   157
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblNetoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   3
            Left            =   -64560
            TabIndex        =   156
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblVentaD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   3
            Left            =   -64560
            TabIndex        =   155
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo de Insumos  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   69
            Left            =   -74880
            TabIndex        =   151
            Top             =   540
            Width           =   1770
         End
         Begin VB.Label lblInsumo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -73080
            TabIndex        =   150
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "&Gastos Generales  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   68
            Left            =   -74880
            TabIndex        =   149
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Costo M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   67
            Left            =   -74880
            TabIndex        =   148
            Top             =   1620
            Width           =   1770
         End
         Begin VB.Label lblCosto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Index           =   2
            Left            =   -73080
            TabIndex        =   147
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mano de &Obra M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   66
            Left            =   -74880
            TabIndex        =   146
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblGasto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -73080
            TabIndex        =   145
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblMObra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -73080
            TabIndex        =   144
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   65
            Left            =   -70920
            TabIndex        =   143
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   64
            Left            =   -70920
            TabIndex        =   142
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Utilidad Sugerida M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   -70320
            TabIndex        =   141
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Sugerido M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   62
            Left            =   -70560
            TabIndex        =   140
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta Neto M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   61
            Left            =   -70320
            TabIndex        =   139
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   60
            Left            =   -69960
            TabIndex        =   138
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label lblUtilidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -68640
            TabIndex        =   137
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -68640
            TabIndex        =   136
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblVentaS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   2
            Left            =   -68640
            TabIndex        =   135
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblNetoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   2
            Left            =   -68640
            TabIndex        =   134
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   59
            Left            =   -66360
            TabIndex        =   133
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblUtilidadReal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -67200
            TabIndex        =   132
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   38
            Left            =   -66360
            TabIndex        =   131
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   -65880
            TabIndex        =   130
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   -65115
            TabIndex        =   129
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   -65115
            TabIndex        =   128
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   -65115
            TabIndex        =   127
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label lblTCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -64560
            TabIndex        =   126
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   -64560
            TabIndex        =   125
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblNetoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   2
            Left            =   -64560
            TabIndex        =   124
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblVentaD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   2
            Left            =   -64560
            TabIndex        =   123
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo de Insumos  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   -74880
            TabIndex        =   119
            Top             =   540
            Width           =   1770
         End
         Begin VB.Label lblInsumo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -73080
            TabIndex        =   118
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "&Gastos Generales  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   -74880
            TabIndex        =   117
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Costo M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   -74880
            TabIndex        =   116
            Top             =   1620
            Width           =   1770
         End
         Begin VB.Label lblCosto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Index           =   1
            Left            =   -73080
            TabIndex        =   115
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mano de &Obra M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   -74880
            TabIndex        =   114
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblGasto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -73080
            TabIndex        =   113
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblMObra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -73080
            TabIndex        =   112
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   29
            Left            =   -70920
            TabIndex        =   111
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   28
            Left            =   -70920
            TabIndex        =   110
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Utilidad Sugerida M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   -70320
            TabIndex        =   109
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Sugerido M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   -70560
            TabIndex        =   108
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta Neto M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   -70320
            TabIndex        =   107
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   -69960
            TabIndex        =   106
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label lblUtilidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -68640
            TabIndex        =   105
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -68640
            TabIndex        =   104
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblVentaS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   1
            Left            =   -68640
            TabIndex        =   103
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblNetoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   1
            Left            =   -68640
            TabIndex        =   102
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   23
            Left            =   -66360
            TabIndex        =   101
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblUtilidadReal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -67200
            TabIndex        =   100
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   22
            Left            =   -66360
            TabIndex        =   99
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   -65880
            TabIndex        =   98
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   -65115
            TabIndex        =   97
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   -65115
            TabIndex        =   96
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   -65115
            TabIndex        =   95
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label lblTCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -64560
            TabIndex        =   94
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblSugeridoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   -64560
            TabIndex        =   93
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblNetoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   1
            Left            =   -64560
            TabIndex        =   92
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblVentaD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   1
            Left            =   -64560
            TabIndex        =   91
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblVentaD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   87
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblNetoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   86
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblSugeridoD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   85
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblTCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   10440
            TabIndex        =   84
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   9885
            TabIndex        =   83
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   9885
            TabIndex        =   82
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UK$ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   9885
            TabIndex        =   81
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   9120
            TabIndex        =   80
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   8640
            TabIndex        =   79
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label lblUtilidadReal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   7800
            TabIndex        =   78
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   8640
            TabIndex        =   77
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblNetoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   0
            Left            =   6360
            TabIndex        =   75
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblVentaS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Index           =   0
            Left            =   6360
            TabIndex        =   74
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label lblSugeridoS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   6360
            TabIndex        =   73
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label lblUtilidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   6360
            TabIndex        =   72
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   5040
            TabIndex        =   71
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Venta Neto M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   4680
            TabIndex        =   70
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Sugerido M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   4440
            TabIndex        =   69
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Utilidad Sugerida M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   4680
            TabIndex        =   68
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   4080
            TabIndex        =   67
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4080
            TabIndex        =   66
            Top             =   840
            Width           =   210
         End
         Begin VB.Label lblMObra 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   63
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label lblGasto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   62
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mano de &Obra M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   61
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblCosto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   60
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Costo M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   59
            Top             =   1620
            Width           =   1770
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "&Gastos Generales  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   58
            Top             =   900
            Width           =   1770
         End
         Begin VB.Label lblInsumo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   57
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo de Insumos  M/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   56
            Top             =   540
            Width           =   1770
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   58
            Left            =   -74865
            TabIndex        =   55
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   57
            Left            =   -74475
            TabIndex        =   54
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   56
            Left            =   -74865
            TabIndex        =   53
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Index           =   55
            Left            =   -74715
            TabIndex        =   52
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   54
            Left            =   -74430
            TabIndex        =   51
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   53
            Left            =   -74865
            TabIndex        =   50
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   52
            Left            =   -74475
            TabIndex        =   49
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   51
            Left            =   -74865
            TabIndex        =   48
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Index           =   50
            Left            =   -74715
            TabIndex        =   47
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   49
            Left            =   -74430
            TabIndex        =   46
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   48
            Left            =   -74865
            TabIndex        =   45
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   47
            Left            =   -74475
            TabIndex        =   44
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   46
            Left            =   -74865
            TabIndex        =   43
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Index           =   45
            Left            =   -74715
            TabIndex        =   42
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   44
            Left            =   -74430
            TabIndex        =   41
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio Costo :"
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
            Index           =   43
            Left            =   -74865
            TabIndex        =   40
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Insumos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   42
            Left            =   -74475
            TabIndex        =   39
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Grals :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   41
            Left            =   -74865
            TabIndex        =   38
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mano Obra :"
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
            Index           =   40
            Left            =   -74715
            TabIndex        =   37
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   39
            Left            =   -74430
            TabIndex        =   36
            Top             =   2565
            Width           =   780
         End
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1470
         TabIndex        =   9
         Top             =   315
         Width           =   7920
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Receta Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   8
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   1470
         TabIndex        =   7
         Top             =   750
         Width           =   2820
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Su&b Almacén :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   0
         Top             =   825
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmRecetaVentaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset
Dim RsCabecera As Recordset
Dim nNeto As Double
Dim nNeto1 As Double
Dim nNeto2 As Double
Dim nNeto3 As Double
Dim nNeto4 As Double
Dim nPorcentaje As Double

Private Sub cmdOpcion_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
   Me.Caption = " Recetas de Venta "
   
   LabelCanales
   
   Call ConfGrilla(9, grdGrilla, "Codigo", 2, "tCodigoProducto", 900, 2, 0, "", _
                                 "Articulo", 2, "Producto", 2500, 0, 0, "", _
                                 "Und.Entrada", 2, "UnidadEntrada", 1500, 0, 0, "", _
                                 "Precio", 2, "PrecioPromedio", 1100, 1, 0, "##0.000", _
                                 "Factor", 2, "Factor", 900, 1, 0, "##0.00", _
                                 "Und.Salida", 2, "UnidadSalida", 1500, 0, 0, "", _
                                 "Cantidad", 2, "nCantidad", 1200, 1, 0, "##0.000", _
                                 "Sub Total", 2, "SubTotal", 1200, 1, 0, "##0.000", _
                                 "Desc", 2, "lDescargo", 700, 2, 4, "")
                            
    Isql = "SELECT DISTINCT dbo.DRECETAVENTA.*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.TPRODUCTO.nPrecioPromedio AS PrecioPromedio, dbo.TPRODUCTO.nFactor AS Factor, vUnidadMedida_1.DESCRIPCION AS UnidadEntrada, vUnidadMedida_2.DESCRIPCION AS UnidadSalida, " & _
           "dbo.TPRODUCTO.nPrecioPromedio / dbo.TPRODUCTO.nFactor AS PrecioSalida, dbo.TPRODUCTO.nPrecioPromedio / dbo.TPRODUCTO.nFactor * dbo.DRECETAVENTA.nCantidad AS SubTotal " & _
           "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DRECETAVENTA ON dbo.TPRODUCTO.tCodigoProducto = dbo.DRECETAVENTA.tCodigoProducto LEFT OUTER JOIN dbo.vUnidadMedida vUnidadMedida_2 ON dbo.TPRODUCTO.tUnidadSalida = vUnidadMedida_2.CODIGO LEFT OUTER JOIN " & _
           "dbo.vUnidadMedida vUnidadMedida_1 ON dbo.TPRODUCTO.tUnidadEntrada = vUnidadMedida_1.CODIGO " & _
           "where tRecetaVenta='" & sCodigo & "' and tLocal='" & sLocal & "'"
           
    Set RsGrilla = Lib.OpenRecordset(Isql, CnAlmacen)
    Set grdGrilla.DataSource = RsGrilla
    
    Isql = "select * from vRECETAVENTA where tRecetaVenta='" & sCodigo & "' and tLocal='" & sLocal & "'"
    Set RsCabecera = Lib.OpenRecordset(Isql, CnAlmacen)
    If RsCabecera.RecordCount <> 0 Then
       Asignar
    End If
    Screen.MousePointer = vbDefault
End Sub

Sub Asignar()
   Dim nTotalCosto As Double, nSugerido As Double
   Dim nTotalCosto1 As Double, nSugerido1 As Double
   Dim nTotalCosto2 As Double, nSugerido2 As Double
   Dim nTotalCosto3 As Double, nSugerido3 As Double
   Dim nTotalCosto4 As Double, nSugerido4 As Double
   Screen.MousePointer = vbHourglass
   
   TabCostos.TabCaption(0) = IIf(sBoton1 = "", "n/d", sBoton1)
   TabCostos.TabCaption(1) = IIf(sBoton2 = "", "n/d", sBoton2)
   TabCostos.TabCaption(2) = IIf(sBoton3 = "", "n/d", sBoton3)
   TabCostos.TabCaption(3) = IIf(sBoton4 = "", "n/d", sBoton4)
   TabCostos.TabCaption(4) = IIf(sBoton5 = "", "n/d", sBoton5)
   
   With RsCabecera
        Frame.Caption = " " & IIf(IsNull(!tRecetaventa), "", !tRecetaventa) & " "
        lblArea.Caption = IIf(IsNull(!Area), "", !Area)
        lblDescripcion.Caption = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        chkActivo.Value = IIf(!lActivo, 1, 0)
        txtObservacion.Text = IIf(IsNull(!tObservacion), "", !tObservacion)

        Select Case pais
            Case "001" 'Bolivia
                With frmProducto.RsCabecera
                     nPorcentaje = ((IIf(!lImpuesto1, nPorcentaje1, 0) + IIf(!lImpuesto2, nPorcentaje2, 0) + IIf(!lImpuesto3, nPorcentaje3, 0)) / 100)
                     nNeto = IIf(IsNull(!nPrecioVenta), 0, (!nPrecioVenta) - (!nPrecioVenta * nPorcentaje))
                     nNeto1 = IIf(IsNull(!nPrecioDelivery), 0, (!nPrecioDelivery) - (!nPrecioDelivery * nPorcentaje))
                     nNeto2 = IIf(IsNull(!nPrecioLlevar), 0, (!nPrecioLlevar) - (!nPrecioLlevar * nPorcentaje))
                     nNeto3 = IIf(IsNull(!nPrecioCanal4), 0, (!nPrecioCanal4) - (!nPrecioCanal4 * nPorcentaje))
                     nNeto4 = IIf(IsNull(!nPrecioCanal5), 0, (!nPrecioCanal5) - (!nPrecioCanal5 * nPorcentaje))
                     
                     lblNetoS(0).Caption = Format(nNeto, "###,##0.00")
                     lblNetoS(1).Caption = Format(nNeto1, "###,##0.00")
                     lblNetoS(2).Caption = Format(nNeto2, "###,##0.00")
                     lblNetoS(3).Caption = Format(nNeto3, "###,##0.00")
                     lblNetoS(4).Caption = Format(nNeto4, "###,##0.00")
                     
                     lblVentaS(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta), "###,##0.00")
                     lblVentaS(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery), "###,##0.00")
                     lblVentaS(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar), "###,##0.00")
                     lblVentaS(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4), "###,##0.00")
                     lblVentaS(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5), "###,##0.00")
                     
                     lblNetoD(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, ((!nPrecioVenta) - (!nPrecioVenta * nPorcentaje)) / nTC), "###,##0.00")
                     lblNetoD(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, ((!nPrecioDelivery) - (!nPrecioDelivery * nPorcentaje)) / nTC), "###,##0.00")
                     lblNetoD(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, ((!nPrecioLlevar) - (!nPrecioLlevar * nPorcentaje)) / nTC), "###,##0.00")
                     lblNetoD(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, ((!nPrecioCanal4) - (!nPrecioCanal4 * nPorcentaje)) / nTC), "###,##0.00")
                     lblNetoD(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, ((!nPrecioCanal5) - (!nPrecioCanal5 * nPorcentaje)) / nTC), "###,##0.00")

                     
                     lblVentaD(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta / nTC), "###,##0.00")
                     lblVentaD(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery / nTC), "###,##0.00")
                     lblVentaD(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar / nTC), "###,##0.00")
                     lblVentaD(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4 / nTC), "###,##0.00")
                     lblVentaD(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5 / nTC), "###,##0.00")
                End With

            Case Else 'Peru, Ecuador
                With frmProducto.RsCabecera
                     nPorcentaje = 1 + ((IIf(!lImpuesto1, nPorcentaje1, 0) + IIf(!lImpuesto2, nPorcentaje2, 0) + IIf(!lImpuesto3, nPorcentaje3, 0)) / 100)
                     nNeto = IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta / nPorcentaje)
                     nNeto1 = IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery / nPorcentaje)
                     nNeto2 = IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar / nPorcentaje)
                     nNeto3 = IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4 / nPorcentaje)
                     nNeto4 = IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5 / nPorcentaje)
                     
                     lblNetoS(0).Caption = Format(nNeto, "###,##0.00")
                     lblNetoS(1).Caption = Format(nNeto1, "###,##0.00")
                     lblNetoS(2).Caption = Format(nNeto2, "###,##0.00")
                     lblNetoS(3).Caption = Format(nNeto3, "###,##0.00")
                     lblNetoS(4).Caption = Format(nNeto4, "###,##0.00")
                                         
                     lblVentaS(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta), "###,##0.00")
                     lblVentaS(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery), "###,##0.00")
                     lblVentaS(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar), "###,##0.00")
                     lblVentaS(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4), "###,##0.00")
                     lblVentaS(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5), "###,##0.00")
                     
                     lblNetoD(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta / nPorcentaje / nTC), "###,##0.00")
                     lblNetoD(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery / nPorcentaje / nTC), "###,##0.00")
                     lblNetoD(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar / nPorcentaje / nTC), "###,##0.00")
                     lblNetoD(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4 / nPorcentaje / nTC), "###,##0.00")
                     lblNetoD(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5 / nPorcentaje / nTC), "###,##0.00")

                     lblVentaD(0).Caption = Format(IIf(IsNull(!nPrecioVenta), 0, !nPrecioVenta / nTC), "###,##0.00")
                     lblVentaD(1).Caption = Format(IIf(IsNull(!nPrecioDelivery), 0, !nPrecioDelivery / nTC), "###,##0.00")
                     lblVentaD(2).Caption = Format(IIf(IsNull(!nPrecioLlevar), 0, !nPrecioLlevar / nTC), "###,##0.00")
                     lblVentaD(3).Caption = Format(IIf(IsNull(!nPrecioCanal4), 0, !nPrecioCanal4 / nTC), "###,##0.00")
                     lblVentaD(4).Caption = Format(IIf(IsNull(!nPrecioCanal5), 0, !nPrecioCanal5 / nTC), "###,##0.00")
                End With


        End Select
        'CANAL 1
        txtGasto(0).Text = Format(IIf(IsNull(!nPorcentajeGastoGeneral), 0, !nPorcentajeGastoGeneral), "##0.00")
        txtMObra(0).Text = Format(IIf(IsNull(!nPorcentajeManoObra), 0, !nPorcentajeManoObra), "##0.00")
        txtUtilidad(0).Text = Format(IIf(IsNull(!nPorcentajeUtilidad), 0, !nPorcentajeUtilidad), "##0.00")

        nTotalCosto = IIf(IsNull(!nInsumo), 0, !nInsumo) + IIf(IsNull(!nManoObra), 0, !nManoObra) + IIf(IsNull(!nGastoGeneral), 0, !nGastoGeneral)
        nSugerido = nTotalCosto + (nTotalCosto * IIf(IsNull(!nPorcentajeUtilidad), 0, !nPorcentajeUtilidad)) / 100

        lblInsumo(0).Caption = Format(IIf(IsNull(!nInsumo), 0, !nInsumo), "###,##0.000")
        lblGasto(0).Caption = Format(IIf(IsNull(!nGastoGeneral), 0, !nGastoGeneral), "###,##0.000")
        lblMObra(0).Caption = Format(IIf(IsNull(!nManoObra), 0, !nManoObra), "###,##0.000")
        lblUtilidad(0).Caption = Format((nTotalCosto * IIf(IsNull(!nPorcentajeUtilidad), 0, !nPorcentajeUtilidad)), "###,##0.000")
        lblCosto(0).Caption = Format(nTotalCosto, "###,##0.000")

        If nNeto > 0 Then
           Dim nTotalC As Double
           nTotalC = !nInsumo + !nGastoGeneral + !nManoObra
           lblUtilidadReal(0).Caption = Format(((nNeto - IIf(nTotalC = 0, 0, nTotalC)) * 100 / nNeto), "##0.00")
        Else
           lblUtilidadReal(0).Caption = "0.00"
        End If
        lblSugeridoS(0).Caption = Format(nSugerido, "###,##0.000")
        lblSugeridoD(0).Caption = Format(nSugerido / nTC, "###,##0.000")
        
        'CANAL 2
        txtGasto(1).Text = Format(IIf(IsNull(!nPorcentajeGastoGeneral2), 0, !nPorcentajeGastoGeneral2), "##0.00")
        txtMObra(1).Text = Format(IIf(IsNull(!nPorcentajeManoObra2), 0, !nPorcentajeManoObra2), "##0.00")
        txtUtilidad(1).Text = Format(IIf(IsNull(!nPorcentajeUtilidad2), 0, !nPorcentajeUtilidad2), "##0.00")

        nTotalCosto1 = IIf(IsNull(!nInsumo2), 0, !nInsumo2) + IIf(IsNull(!nManoObra2), 0, !nManoObra2) + IIf(IsNull(!nGastoGeneral2), 0, !nGastoGeneral2)
        nSugerido1 = nTotalCosto1 + (nTotalCosto1 * IIf(IsNull(!nPorcentajeUtilidad2), 0, !nPorcentajeUtilidad2)) / 100

        lblInsumo(1).Caption = Format(IIf(IsNull(!nInsumo2), 0, !nInsumo2), "###,##0.000")
        lblGasto(1).Caption = Format(IIf(IsNull(!nGastoGeneral2), 0, !nGastoGeneral2), "###,##0.000")
        lblMObra(1).Caption = Format(IIf(IsNull(!nManoObra2), 0, !nManoObra2), "###,##0.000")
        lblUtilidad(1).Caption = Format((nTotalCosto1 * IIf(IsNull(!nPorcentajeUtilidad2), 0, !nPorcentajeUtilidad2)), "###,##0.000")
        lblCosto(1).Caption = Format(nTotalCosto1, "###,##0.000")

        If nNeto1 > 0 Then
           Dim nTotalC1 As Double
           nTotalC1 = !nInsumo2 + !nGastoGeneral2 + !nManoObra2
           lblUtilidadReal(1).Caption = Format(((nNeto1 - IIf(nTotalC1 = 0, 0, nTotalC1)) * 100 / nNeto1), "##0.00")
        Else
           lblUtilidadReal(1).Caption = "0.00"
        End If
        lblSugeridoS(1).Caption = Format(nSugerido1, "###,##0.000")
        lblSugeridoD(1).Caption = Format(nSugerido1 / nTC, "###,##0.000")
        
        'CANAL 3
        txtGasto(2).Text = Format(IIf(IsNull(!nPorcentajeGastoGeneral3), 0, !nPorcentajeGastoGeneral3), "##0.00")
        txtMObra(2).Text = Format(IIf(IsNull(!nPorcentajeManoObra3), 0, !nPorcentajeManoObra3), "##0.00")
        txtUtilidad(2).Text = Format(IIf(IsNull(!nPorcentajeUtilidad3), 0, !nPorcentajeUtilidad3), "##0.00")

        nTotalCosto2 = IIf(IsNull(!nInsumo3), 0, !nInsumo3) + IIf(IsNull(!nManoObra3), 0, !nManoObra3) + IIf(IsNull(!nGastoGeneral3), 0, !nGastoGeneral3)
        nSugerido2 = nTotalCosto2 + (nTotalCosto2 * IIf(IsNull(!nPorcentajeUtilidad3), 0, !nPorcentajeUtilidad3)) / 100

        lblInsumo(2).Caption = Format(IIf(IsNull(!nInsumo3), 0, !nInsumo3), "###,##0.000")
        lblGasto(2).Caption = Format(IIf(IsNull(!nGastoGeneral3), 0, !nGastoGeneral3), "###,##0.000")
        lblMObra(2).Caption = Format(IIf(IsNull(!nManoObra3), 0, !nManoObra3), "###,##0.000")
        lblUtilidad(2).Caption = Format((nTotalCosto2 * IIf(IsNull(!nPorcentajeUtilidad3), 0, !nPorcentajeUtilidad3)), "###,##0.000")
        lblCosto(2).Caption = Format(nTotalCosto2, "###,##0.000")

        If nNeto2 > 0 Then
           Dim nTotalC2 As Double
           nTotalC2 = !nInsumo3 + !nGastoGeneral3 + !nManoObra3
           lblUtilidadReal(2).Caption = Format(((nNeto2 - IIf(nTotalC2 = 0, 0, nTotalC2)) * 100 / nNeto2), "##0.00")
        Else
           lblUtilidadReal(2).Caption = "0.00"
        End If
        lblSugeridoS(2).Caption = Format(nSugerido2, "###,##0.000")
        lblSugeridoD(2).Caption = Format(nSugerido2 / nTC, "###,##0.000")
        
        'CANAL 4
        txtGasto(3).Text = Format(IIf(IsNull(!nPorcentajeGastoGeneral4), 0, !nPorcentajeGastoGeneral4), "##0.00")
        txtMObra(3).Text = Format(IIf(IsNull(!nPorcentajeManoObra4), 0, !nPorcentajeManoObra4), "##0.00")
        txtUtilidad(3).Text = Format(IIf(IsNull(!nPorcentajeUtilidad4), 0, !nPorcentajeUtilidad4), "##0.00")

        nTotalCosto3 = IIf(IsNull(!nInsumo4), 0, !nInsumo4) + IIf(IsNull(!nManoObra4), 0, !nManoObra4) + IIf(IsNull(!nGastoGeneral4), 0, !nGastoGeneral4)
        nSugerido3 = nTotalCosto3 + (nTotalCosto3 * IIf(IsNull(!nPorcentajeUtilidad4), 0, !nPorcentajeUtilidad4)) / 100

        lblInsumo(3).Caption = Format(IIf(IsNull(!nInsumo4), 0, !nInsumo4), "###,##0.000")
        lblGasto(3).Caption = Format(IIf(IsNull(!nGastoGeneral4), 0, !nGastoGeneral4), "###,##0.000")
        lblMObra(3).Caption = Format(IIf(IsNull(!nManoObra4), 0, !nManoObra4), "###,##0.000")
        lblUtilidad(3).Caption = Format((nTotalCosto3 * IIf(IsNull(!nPorcentajeUtilidad4), 0, !nPorcentajeUtilidad4)), "###,##0.000")
        lblCosto(3).Caption = Format(nTotalCosto3, "###,##0.000")

        If nNeto3 > 0 Then
           Dim nTotalC3 As Double
           nTotalC3 = !nInsumo4 + !nGastoGeneral4 + !nManoObra4
           lblUtilidadReal(3).Caption = Format(((nNeto3 - IIf(nTotalC3 = 0, 0, nTotalC3)) * 100 / nNeto3), "##0.00")
        Else
           lblUtilidadReal(3).Caption = "0.00"
        End If
        lblSugeridoS(3).Caption = Format(nSugerido3, "###,##0.000")
        lblSugeridoD(3).Caption = Format(nSugerido3 / nTC, "###,##0.000")


        'CANAL 5
        txtGasto(4).Text = Format(IIf(IsNull(!nPorcentajeGastoGeneral5), 0, !nPorcentajeGastoGeneral5), "##0.00")
        txtMObra(4).Text = Format(IIf(IsNull(!nPorcentajeManoObra5), 0, !nPorcentajeManoObra5), "##0.00")
        txtUtilidad(4).Text = Format(IIf(IsNull(!nPorcentajeUtilidad5), 0, !nPorcentajeUtilidad5), "##0.00")

        nTotalCosto4 = IIf(IsNull(!nInsumo5), 0, !nInsumo5) + IIf(IsNull(!nManoObra5), 0, !nManoObra5) + IIf(IsNull(!nGastoGeneral5), 0, !nGastoGeneral5)
        nSugerido4 = nTotalCosto4 + (nTotalCosto4 * IIf(IsNull(!nPorcentajeUtilidad5), 0, !nPorcentajeUtilidad5)) / 100

        lblInsumo(4).Caption = Format(IIf(IsNull(!nInsumo5), 0, !nInsumo5), "###,##0.000")
        lblGasto(4).Caption = Format(IIf(IsNull(!nGastoGeneral5), 0, !nGastoGeneral5), "###,##0.000")
        lblMObra(4).Caption = Format(IIf(IsNull(!nManoObra5), 0, !nManoObra5), "###,##0.000")
        lblUtilidad(4).Caption = Format((nTotalCosto4 * IIf(IsNull(!nPorcentajeUtilidad5), 0, !nPorcentajeUtilidad5)), "###,##0.000")
        lblCosto(4).Caption = Format(nTotalCosto4, "###,##0.000")

        If nNeto4 > 0 Then
           Dim nTotalC4 As Double
           nTotalC4 = !nInsumo5 + !nGastoGeneral5 + !nManoObra5
           lblUtilidadReal(4).Caption = Format(((nNeto4 - IIf(nTotalC4 = 0, 0, nTotalC4)) * 100 / nNeto4), "##0.00")
        Else
           lblUtilidadReal(4).Caption = "0.00"
        End If
        lblSugeridoS(4).Caption = Format(nSugerido4, "###,##0.000")
        lblSugeridoD(4).Caption = Format(nSugerido4 / nTC, "###,##0.000")
        
  End With

  Screen.MousePointer = vbDefault
End Sub

Sub LabelCanales()
  'CANAL1
   Label(9).Caption = "Insumos " & sMonN & " : "
   Label(10).Caption = "&Gastos Generales " & sMonN & " : "
   Label(11).Caption = "Mano de &Obra " & sMonN & " : "
   Label(12).Caption = "Costo Total " & sMonN & " : "
   
   Label(13).Caption = "&Utilidad Sugerida " & sMonN & " : "
   Label(14).Caption = "Valor Neto Sugerido " & sMonN & " : "
   Label(15).Caption = "Valor Neto " & sMonN & " : "
   Label(16).Caption = "Valor Venta " & sMonN & " : "
      
   Label(17).Caption = "Valor Neto Sugerido " & sMonE & " : "
   Label(18).Caption = "Valor Neto " & sMonE & " : "
   Label(19).Caption = "Valor Venta " & sMonE & " : "
   
   'CANAL2
   Label(33).Caption = "Insumos " & sMonN & " : "
   Label(32).Caption = "&Gastos Generales " & sMonN & " : "
   Label(30).Caption = "Mano de &Obra " & sMonN & " : "
   Label(31).Caption = "Costo Total " & sMonN & " : "
   
   Label(27).Caption = "&Utilidad Sugerida " & sMonN & " : "
   Label(26).Caption = "Valor Neto Sugerido " & sMonN & " : "
   Label(25).Caption = "Valor Neto " & sMonN & " : "
   Label(24).Caption = "Valor Venta " & sMonN & " : "
      
   Label(20).Caption = "Valor Neto Sugerido " & sMonE & " : "
   Label(7).Caption = "Valor Neto " & sMonE & " : "
   Label(6).Caption = "Valor Venta " & sMonE & " : "
   
   'CANAL3
   Label(69).Caption = "Insumos " & sMonN & " : "
   Label(68).Caption = "&Gastos Generales " & sMonN & " : "
   Label(66).Caption = "Mano de &Obra " & sMonN & " : "
   Label(67).Caption = "Costo Total " & sMonN & " : "
   
   Label(63).Caption = "&Utilidad Sugerida " & sMonN & " : "
   Label(62).Caption = "Valor Neto Sugerido " & sMonN & " : "
   Label(61).Caption = "Valor Neto " & sMonN & " : "
   Label(60).Caption = "Valor Venta " & sMonN & " : "
      
   Label(36).Caption = "Valor Neto Sugerido " & sMonE & " : "
   Label(35).Caption = "Valor Neto " & sMonE & " : "
   Label(34).Caption = "Valor Venta " & sMonE & " : "
   
   'CANAL4
   Label(85).Caption = "Insumos " & sMonN & " : "
   Label(84).Caption = "&Gastos Generales " & sMonN & " : "
   Label(82).Caption = "Mano de &Obra " & sMonN & " : "
   Label(83).Caption = "Costo Total " & sMonN & " : "
   
   Label(79).Caption = "&Utilidad Sugerida " & sMonN & " : "
   Label(78).Caption = "Valor Neto Sugerido " & sMonN & " : "
   Label(77).Caption = "Valor Neto " & sMonN & " : "
   Label(76).Caption = "Valor Venta " & sMonN & " : "
      
   Label(72).Caption = "Valor Neto Sugerido " & sMonE & " : "
   Label(71).Caption = "Valor Neto " & sMonE & " : "
   Label(70).Caption = "Valor Venta " & sMonE & " : "
   
   'CANAL5
   Label(101).Caption = "Insumos " & sMonN & " : "
   Label(100).Caption = "&Gastos Generales " & sMonN & " : "
   Label(98).Caption = "Mano de &Obra " & sMonN & " : "
   Label(99).Caption = "Costo Total " & sMonN & " : "
   
   Label(95).Caption = "&Utilidad Sugerida " & sMonN & " : "
   Label(94).Caption = "Valor Neto Sugerido " & sMonN & " : "
   Label(93).Caption = "Valor Neto " & sMonN & " : "
   Label(92).Caption = "Valor Venta " & sMonN & " : "
      
   Label(88).Caption = "Valor Neto Sugerido " & sMonE & " : "
   Label(87).Caption = "Valor Neto " & sMonE & " : "
   Label(86).Caption = "Valor Venta " & sMonE & " : "
   
   
   lblTCambio(0).Caption = Format(nTC, "#,##0.000")
   lblTCambio(1).Caption = Format(nTC, "#,##0.000")
   lblTCambio(2).Caption = Format(nTC, "#,##0.000")
   lblTCambio(3).Caption = Format(nTC, "#,##0.000")
   lblTCambio(4).Caption = Format(nTC, "#,##0.000")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsCabecera = Nothing
   Set RsGrilla = Nothing
   Set frmRecetaVentaDetalle = Nothing
End Sub


