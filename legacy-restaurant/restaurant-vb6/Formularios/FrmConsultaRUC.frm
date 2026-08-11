VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConsultaRUC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar RUC"
   ClientHeight    =   9840
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   15705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnVer 
      Caption         =   "Ver"
      Height          =   615
      Left            =   13800
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Mismo RUC, diferente Razon Social"
      TabPicture(0)   =   "FrmConsultaRUC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblTotal1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblTotal2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdGrilla3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "grdGrilla"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdMovimiento(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdMovimiento(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMovimiento(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdMovimiento(11)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOpcion(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "RUC no Encontrado"
      TabPicture(1)   =   "FrmConsultaRUC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdGrilla2"
      Tab(1).Control(1)=   "LblTotal3"
      Tab(1).Control(2)=   "Label3"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ejecutar"
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
         Index           =   9
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   11
         Left            =   9720
         Picture         =   "FrmConsultaRUC.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3050
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   9
         Left            =   9720
         Picture         =   "FrmConsultaRUC.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1665
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   8
         Left            =   9720
         Picture         =   "FrmConsultaRUC.frx":0B4C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   10
         Left            =   9720
         Picture         =   "FrmConsultaRUC.frx":10D6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2355
         Width           =   1260
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7995
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   14102
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Mismo RUC con Razon Social Diferente"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(40)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(41)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(44)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(45)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(46)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(51)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(57)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(58)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(61)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(65)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(69)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(70)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(73)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(74)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(75)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(76)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(77)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(78)  =   ":id=13,.fontname=Arial Black"
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla2 
         Height          =   7995
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14102
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "RUC no Encontrado"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(40)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(41)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(44)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(45)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(46)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(51)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(57)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(58)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(61)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(65)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(69)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(70)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(73)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(74)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(75)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(76)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(77)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(78)  =   ":id=13,.fontname=Arial Black"
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla3 
         Height          =   7995
         Left            =   11040
         TabIndex        =   6
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   14102
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "RUC Por Cambiar"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(40)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(41)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(44)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(45)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(46)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(51)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(57)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(58)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(61)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(65)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(69)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(70)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(73)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(74)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(75)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(76)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(77)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(78)  =   ":id=13,.fontname=Arial Black"
      End
      Begin VB.Label LblTotal3 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   -73200
         TabIndex        =   18
         Top             =   8520
         Width           =   1095
      End
      Begin VB.Label LblTotal2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   10920
         TabIndex        =   17
         Top             =   8520
         Width           =   1095
      End
      Begin VB.Label LblTotal1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   8520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad de Registro:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad de Registro:"
         Height          =   255
         Left            =   9360
         TabIndex        =   14
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de Registro:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   8520
         Width           =   1575
      End
   End
   Begin VB.CommandButton BtnConsulta 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   11160
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   5000
      Scrolling       =   1
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Registro Nº 1 de 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmConsultaRUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ix As Integer
Dim RsConsulta As Recordset
Dim RsLista As Recordset
Dim RsLista2 As Recordset
Dim RsListaNo As Recordset
Dim cont As Double

Private Sub BtnConsulta_Click()


    If MsgBox("Seguro de realizar el proceso???, Recordar que el proceso tomara tiempo de acuerdo a la cantidad de Datos que tenga por Validar", vbYesNo + vbQuestion, sMensaje) = vbNo Then
        Exit Sub
    End If
    
On Error GoTo fin

'-- 0 por valida
'-- 1 validado
'-- 2 x cambiar
'-- 3 actualizado

    Screen.MousePointer = vbHourglass
    
 
    
    
    
    
    
    
    pbProgress.Value = 0

    Isql = "select * from vCliente  where tTipoIdentidad ='02' and lActivo=1  and isnull(revisado,0) = 0 order by Codigo"
    Set RsConsulta = Lib.OpenRecordset(Isql, Cn)
    pbProgress.Max = RsConsulta.RecordCount


    For ix = 1 To RsConsulta.RecordCount
        DoEvents
        Dim xDat As String
        Dim xWml As New XMLHTTP
         cont = cont + 1

        xWml.Open "POST", RutaConsultaRuc & "/" & Trim(RsConsulta!tIdentidad), False
        xWml.send
        If xWml.Status = 200 Then
            xDat = Trim(xWml.responseText)
            xDat = Replace(xDat, vbNewLine, " ")
            xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
            xDat = Replace(xDat, Chr(34), "")

            Dim MATRIZ() As String
            MATRIZ = Split(xDat, vbNewLine)

            MATRIZ(0) = Replace(MATRIZ(0), "status:", "")
            If (MATRIZ(0) = 1) Then
                MATRIZ(2) = Replace(MATRIZ(2), "ruc:", "")
                MATRIZ(3) = Replace(MATRIZ(3), "razonSocial:", "")
                MATRIZ(4) = Replace(MATRIZ(4), "direccion:", "")
                MATRIZ(5) = Replace(MATRIZ(5), "ubigeo:", "")

                usp_ActualizarCabPeDoc "3", Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn), MATRIZ(3), MATRIZ(4), MATRIZ(5), RsConsulta!tIdentidad, ""

    '            Me.txtUrbanizacion.Text = Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn)
    '            Me.txtDetallado.Text = MATRIZ(3)
    '            'txtRazonSocial.Text = MATRIZ(0)
    '            Me.txtDireccion.Text = MATRIZ(4)
    '            Me.txtCodigoUbigeo.Text = MATRIZ(5)
    '        Else
                'Call verificRuc
            End If
    '    ElseIf xWml.Status = 500 Then
    '        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    '    ElseIf xWml.Status = 12007 Then
    '        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    '    Else
    '        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
        End If

        lblProgress.Caption = "Procesando Registro " & cont & " de " & RsConsulta.RecordCount
        pbProgress.Value = (cont * RsConsulta.RecordCount) / RsConsulta.RecordCount


        Set xWml = Nothing
        Screen.MousePointer = vbDefault

        RsConsulta.MoveNext
     Next ix
     cont = 0
   ' Exit Sub
    'pbProgress.Value = 100
    lblProgress.Caption = "Proceso terminado..."
    MsgBox "Terminado", vbInformation, sMensaje
    
    ActualizaRUC
    CargarLista
    CargarListaRUCnoExiste
    CargarLista2
fin:
    'MsgBox "Mensaje: " & error
'   MsgBox "Terminadooooooooo", vbInformation, sMensaje
'    Screen.MousePointer = vbDefault
'    ActualizaRUC
'    CargarLista
'    CargarListaRUCnoExiste

End Sub
Sub CargarLista()
'Isql = "select t1.tIdentidad, t1.Descripcion, t2.xruc , t2.xnombre from vCliente t1 left join ruc_temp t2 on t1.tIdentidad = t2.xruc where t1.tTipoIdentidad ='02' and t1.lActivo=1 AND t1.revisado = 1"
'Set RsLista = Lib.OpenRecordset(Isql, Cn)

usp_ActualizarCabPeDoc "7", "1", "8", "", "", "", ""

   Call ConfGrilla(3, grdGrilla, "RUC", 2, "tIdentidad", 1200, 2, 0, "", _
                                  "Razón Social", 2, "Descripcion", 4200, 0, 0, "", _
                                  "Razón Social - SUNAT", 2, "xnombre", 4200, 0, 0, "")
    Set grdGrilla.DataSource = RsLista
    
End Sub
Sub CargarLista2()
'Isql = "select t1.tIdentidad, t1.Descripcion, t2.xruc , t2.xnombre from vCliente t1 left join ruc_temp t2 on t1.tIdentidad = t2.xruc where t1.tTipoIdentidad ='02' and t1.lActivo=1 AND t1.revisado = 2"
'Set RsLista2 = Lib.OpenRecordset(Isql, Cn)
usp_ActualizarCabPeDoc "7", "2", "7", "", "", "", ""
   Call ConfGrilla(2, grdGrilla3, "RUC", 2, "tIdentidad", 1200, 2, 0, "", _
                                  "Razón Social - SUNAT", 2, "xnombre", 4200, 0, 0, "")
    Set grdGrilla3.DataSource = RsLista2
    
End Sub

Sub CargarListaRUCnoExiste()
'Isql = "select t1.tIdentidad, t1.Descripcion, t2.xruc , isnull(t2.xnombre,'RUC NO EXISTE')xnombre from vCliente t1 left join ruc_temp t2 on t1.tIdentidad = t2.xruc where t1.tTipoIdentidad ='02' and t1.lActivo=1 AND isnull(t1.revisado,0) = 0"
'Set RsListaNo = Lib.OpenRecordset(Isql, Cn)
usp_ActualizarCabPeDoc "6", "", "6", "", "", "", ""

   Call ConfGrilla(3, grdGrilla2, "RUC", 2, "tIdentidad", 1200, 2, 0, "", _
                                  "Razón Social", 2, "Descripcion", 5800, 0, 0, "", _
                                  "Razón Social - SUNAT", 2, "xnombre", 7000, 0, 0, "")
    Set grdGrilla2.DataSource = RsListaNo
End Sub
Sub ActualizaRUC()

lblProgress.Caption = "Camparando Información..."

usp_ActualizarCabPeDoc "4", "", "4", "", "", "", ""

lblProgress.Caption = "Proceso terminado..."

MsgBox "Comparación de Datos terminado, se muestra Resultados", vbInformation, sMensaje
 
 
End Sub
Sub Progreso(Etiqueta As String, cont As Double, Registros As Double)
    lblProgress.Caption = Etiqueta
    pbProgress.Value = (cont * 100) / Registros
End Sub
Private Sub BtnVer_Click()
    ActualizaRUC
    CargarLista
    CargarListaRUCnoExiste
    CargarLista2
End Sub
Sub usp_ActualizarCabPeDoc(opcion As String, opcion1 As String, opcion2 As String, opcion3 As String, opcion4 As String, opcion5 As String, opcion6 As String)
Dim oComandox As clsComando
Set oComandox = New clsComando

Set oComandox = New clsComando
If Not oComandox.CreateCmdSp("usp_ActualizarCabPeDoc", Cn) Then
  Set oComandox = Nothing
  Exit Sub
End If

oComandox.CreateParameter "@opcion", adVarChar, adParamInput, 20, opcion
oComandox.CreateParameter "@param1", adVarChar, adParamInput, 300, opcion1
oComandox.CreateParameter "@param2", adVarChar, adParamInput, 300, opcion2
oComandox.CreateParameter "@param3", adVarChar, adParamInput, 300, opcion3
oComandox.CreateParameter "@param4", adVarChar, adParamInput, 300, opcion4
oComandox.CreateParameter "@param5", adVarChar, adParamInput, 300, opcion5
oComandox.CreateParameter "@param6", adVarChar, adParamInput, 300, opcion6

If Not oComandox.GetParamOK Then
  Set oComandox = Nothing
  Exit Sub
End If

If Not oComandox.ExecSP Then
  Set oComandox = Nothing
  Exit Sub
End If

If opcion2 = "6" Then
    Set RsListaNo = oComandox.GetSP()
    LblTotal3.Caption = RsListaNo.RecordCount
    
    
ElseIf opcion2 = "7" Then
    Set RsLista2 = oComandox.GetSP()
    LblTotal2.Caption = RsLista2.RecordCount
    
ElseIf opcion2 = "8" Then
   Set RsLista = oComandox.GetSP()
   LblTotal1.Caption = RsLista.RecordCount
   
End If



End Sub

Private Sub cmdMovimiento_Click(Index As Integer)
  Screen.MousePointer = vbHourglass
    
    Select Case Index
           Case Is = 8  ' Pasar Todas
                Cn.Execute "update TCLIENTE set revisado = 2 where tTipoIdentidad ='02' and lActivo=1  and isnull(revisado,0) = 1"
           
           Case Is = 9  ' Pasar una
                If RsLista.RecordCount <> 0 Then
                  Cn.Execute "update TCLIENTE set revisado =2 where tIdentidad = '" & RsLista!tIdentidad & "'"
                End If
           
           Case Is = 10  ' Quitar una
                If RsLista2.RecordCount <> 0 Then
                   Cn.Execute "update TCLIENTE set revisado =1 where tIdentidad = '" & RsLista2!tIdentidad & "'"
                End If
                
           Case Is = 11 ' Quitar Todas
                Cn.Execute "update TCLIENTE set revisado =1 where tTipoIdentidad ='02' and lActivo=1  and isnull(revisado,0) = 2"
                
    End Select
     
    RsLista.Requery
    RsLista2.Requery


    LblTotal2.Caption = RsLista2.RecordCount
    LblTotal1.Caption = RsLista.RecordCount
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    usp_ActualizarCabPeDoc "5", "", "", "", "", "", ""
    RsLista.Requery
    RsLista2.Requery
    LblTotal2.Caption = RsLista2.RecordCount
    LblTotal1.Caption = RsLista.RecordCount
End Sub

