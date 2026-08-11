VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCombosDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5475
   ClientLeft      =   -570
   ClientTop       =   1875
   ClientWidth     =   11295
   FillColor       =   &H8000000F&
   Icon            =   "frmCombosDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1710
      TabIndex        =   2
      Top             =   0
      Width           =   9555
      Begin VB.CheckBox chkEliminaFijo 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   1710
         TabIndex        =   12
         Top             =   4320
         Width           =   225
      End
      Begin VB.TextBox txtQnCombo 
         Alignment       =   1  'Right Justify
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
         Left            =   3525
         MaxLength       =   10
         TabIndex        =   11
         Text            =   " 0.00"
         Top             =   3960
         Width           =   900
      End
      Begin VB.CheckBox chkFijo 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkUnico 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3525
         TabIndex        =   9
         Top             =   4320
         Width           =   225
      End
      Begin VB.TextBox txtEtiqueta 
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
         Left            =   5760
         MaxLength       =   24
         TabIndex        =   8
         Text            =   " "
         Top             =   4260
         Width           =   1620
      End
      Begin VB.TextBox txtAumPrec 
         Alignment       =   1  'Right Justify
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
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   7
         Text            =   " "
         Top             =   3960
         Width           =   1620
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Eliminar Item"
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
         Left            =   3480
         Picture         =   "frmCombosDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1290
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Agregar Item"
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
         Left            =   2160
         Picture         =   "frmCombosDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1290
      End
      Begin VB.CommandButton cmdCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   9000
         Picture         =   "frmCombosDetalle.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4140
         Width           =   450
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4260
         Width           =   1440
      End
      Begin MSComDlg.CommonDialog cmdColor 
         Left            =   8880
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueOleDBGrid80.TDBGrid grdCombo 
         Height          =   2820
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4974
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
         Caption         =   "Combos"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
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
         _StyleDefs(20)  =   "Splits(0).Style:id=30,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=49,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=31,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=32,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=43,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=45,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=44,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=46,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=47,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=48,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=50,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=51,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=55,.parent=30"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=52,.parent=31,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=53,.parent=32"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=54,.parent=44"
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
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H400040&,.borderColor=&H808000&"
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
      Begin TrueOleDBGrid80.TDBGrid grdCombo1 
         Height          =   2385
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   4207
         _LayoutType     =   4
         _RowHeight      =   21
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
         _StyleDefs(22)  =   "Splits(0).Style:id=25,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=27,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=43,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=54,.parent=25"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=26"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=27"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=43"
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
      Begin VB.Label lblCantidadMinima 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No Eliminar"
         Height          =   195
         Left            =   315
         TabIndex        =   23
         Top             =   4320
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unico"
         Height          =   195
         Left            =   3000
         TabIndex        =   22
         Top             =   4320
         Width           =   450
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cargo automatico"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cantidad a Cargar"
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   3960
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Etiqueta"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   4305
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Aumenta Precio"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   3960
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Máxima :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblCantidad 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mínima :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11235
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4770
      Width           =   11295
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Regresar"
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
         Left            =   10050
         Picture         =   "frmCombosDetalle.frx":1000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   5850
      Left            =   30
      Picture         =   "frmCombosDetalle.frx":10F2
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1620
   End
End
Attribute VB_Name = "frmCombosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tCodigoProducto As String
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim nColumna As Integer
Dim lAsc            As Boolean
Dim rsCombo As New ADODB.Recordset

Private Sub cmdBoton_Click(Index As Integer)
 On Error GoTo cmdCancel
   cmdColor.CancelError = True
   cmdColor.ShowColor
   cmdBoton(Index).backColor = cmdColor.Color
Exit Sub
cmdCancel:
   cmdBoton(Index).backColor = -2147483633
End Sub

Public Sub obtieneCodigoProducto(codigo As String)
    tCodigoProducto = codigo
End Sub

Private Sub cmdCombo_Click(Index As Integer)
    Select Case Index
     Case Is = 0 ' Grabar Cantidad / Automatica
               If rsCombo.RecordCount > 0 Then
                  If val(lblCantidad.Caption) < Calcular("select sum(nCantidad) as Codigo from TCOMBO where tCombo='" & tCodigoProducto & "' and tCodigoProducto<>'" & rsCombo!tCodigoProducto & "' and lFijo=1", Cn) + val(txtQnCombo.Text) Then
                     MsgBox "No puede exceder de la cantidad máxima permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
                     
                  'Cn.Execute "update TCOMBO set nCantidad = " & Val(txtQnCombo.Text) & " , lFijo = " & chkFijo.Value & " where tCombo='" & txtCodigo.Text & "' and tCodigoProducto='" & RsCombo!tCodigoProducto & "'"
                  'Oscar Ortega----------------------------------
                   Cn.Execute "update TCOMBO set nCantidad = " & val(txtQnCombo.Text) & " , lFijo = " & chkFijo.value & " , lUnico = " & chkUnico.value & " , lEliminaFijo = " & chkEliminaFijo.value & ", tEtiqueta = '" & txtEtiqueta.Text & "' , nAumento = " & val(txtAumPrec.Text) & ", nvalor = " & cmdBoton(1).backColor & " where   tCombo='" & tCodigoProducto & "' and tCodigoProducto='" & rsCombo!tCodigoProducto & "'"
                   Cn.Execute "update TCOMBO set lNuevoModificado =1  where   tCombo='" & tCodigoProducto & "' and tCodigoProducto='" & rsCombo!tCodigoProducto & "'"
                  'Fin Oscar Ortega------------------------------
                  rsCombo.Requery
               End If
        Case 1
               If val(lblCantidad.Caption) < Calcular("select sum(nCantidad) as Codigo from TCOMBO where tCombo='" & tCodigoProducto & "' and lFijo=1", Cn) + 1 Then
                  MsgBox "No puede exceder de la cantidad máxima permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lactivo=1 order by Descripcion"
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
               
               If Calcular("select tCodigoProducto as Codigo from TCOMBO where tCombo='" & tCodigoProducto & "' and tCodigoProducto = '" & sCodigo & "'", Cn) > 0 Then
                  MsgBox "Producto ya existente", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                sPasa = tCodigoProducto
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCOMBO", "COMBOS", "01", sUsuario, sPasa, "", _
                   "TCOMBO", "Codigo Combo", tCodigoProducto, "TCODIGOPRODUCTO", "Codigo Producto", sCodigo, "nCantidad", "Cantidad", 1)
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
               
               
               
               
               Dim xpos As Variant
               xpos = frmProducto.RsCabecera.Bookmark
               Cn.Execute "update TPRODUCTO set lCombinacion=1, nCombinacion=" & lblCantidad.Caption & ", lreplica=1 where tCodigoProducto='" & tCodigoProducto & "'"
               
               Cn.Execute "insert into TCOMBO(tCombo, tCodigoProducto, lFijo, nCantidad,nvalor) values ('" & tCodigoProducto & "',  '" & sCodigo & "', 0, 1, -2147483633) "
               frmProducto.RsCabecera.Requery
               frmProducto.RsCabecera.Bookmark = xpos
               rsCombo.Requery
               
        Case 2
               If rsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Producto " & rsCombo!Descripcion & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
      'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCOMBO", "COMBOS", "03", sUsuario, sPasa, "", _
                   "TCOMBO", "Codigo Combo", tCodigoProducto, "TCODIGOPRODUCTO", "Codigo Producto", rsCombo!tCodigoProducto, "nCantidad", "Cantidad", 1)
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                              
               
               'Cambia el Delete
               Cn.Execute "delete from TCOMBO where tCombo='" & tCodigoProducto & "' and tCodigoProducto = '" & rsCombo!tCodigoProducto & "'"
               rsCombo.Requery
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    'Combos
    Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.vProducto.Descripcion, dbo.TCOMBO.nCantidad, dbo.TCOMBO.lFijo, dbo.TCOMBO.lUnico, dbo.TCOMBO.tEtiqueta, dbo.TCOMBO.nAumento,dbo.tcombo.leliminaFijo, ISNULL(dbo.TCOMBO.nValor,0) NVALOR " & _
           "FROM dbo.TCOMBO LEFT OUTER JOIN dbo.vProducto ON dbo.TCOMBO.tCodigoProducto = dbo.vProducto.Codigo " & _
           "where tcombo.tcombo='" & tCodigoProducto & "' " & _
           "Order by tEtiqueta, dbo.vProducto.Descripcion"
    Set rsCombo = Lib.OpenRecordset(Isql, Cn)
   
    'Configuración de la Grilla Combos
    Call ConfGrilla(8, grdCombo, "Codigo", 2, "tCodigoProducto", 820, 2, 0, "", _
                                 "Descripción", 2, "Descripcion", 2600, 0, 0, "", _
                                 "Autom", 2, "lFijo", 780, 2, 4, "", _
                                 "Cant", 2, "nCantidad", 780, 1, 0, "##0.00", _
                                 "Elimi", 2, "lEliminaFijo", 780, 2, 4, "", _
                                 "Unico", 2, "lUnico", 780, 2, 4, "", _
                                 "Etiq", 2, "tEtiqueta", 1400, 0, 0, "", _
                                 "Prec.", 2, "nAumento", 780, 1, 0, "##0.00")
    Set grdCombo.DataSource = rsCombo
End Sub




Private Sub grdCombo_FilterChange()
  On Error GoTo errHandler
    Set cols = grdCombo.Columns
    Dim c As Integer
    c = grdCombo.col
    grdCombo.HoldFields
    rsCombo.Filter = getFilter(cols)
    grdCombo.col = c
    grdCombo.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    rsCombo.Filter = adFilterNone
End Sub

Private Sub grdCombo_HeadClick(ByVal ColIndex As Integer)
  nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       rsCombo.Sort = grdCombo.Columns(ColIndex).DataField & " ASC"
    Else
       rsCombo.Sort = grdCombo.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If rsCombo.RecordCount > 0 Then
      txtQnCombo.Text = Format(rsCombo!nCantidad, "##0.00")
      chkFijo.value = IIf(IsNull(rsCombo!lFijo), 0, Abs(Int(rsCombo!lFijo)))
      chkUnico.value = IIf(IsNull(rsCombo!lUnico), 0, Abs(Int(rsCombo!lUnico)))
      chkEliminaFijo.value = IIf(IsNull(rsCombo!lEliminaFijo), 0, Abs(Int(rsCombo!lEliminaFijo)))
       cmdBoton(1).backColor = IIf(rsCombo!nValor = 0, -2147483633, rsCombo!nValor)
      txtEtiqueta.Text = IIf(IsNull(rsCombo!tEtiqueta), "", rsCombo!tEtiqueta)
      txtAumPrec.Text = Format(rsCombo!nAumento, "##0.00")
   Else
      txtQnCombo.Text = "0.00"
      chkFijo.value = 0
      chkUnico.value = 0
      chkEliminaFijo.value = 0
      txtEtiqueta.Text = ""
      txtAumPrec.Text = Format(0, "##0.00")
   End If
End Sub

 Private Sub chkEliminaFijo_Click()
   If Me.chkEliminaFijo.value = 1 Then
    If Me.chkFijo.value = 0 Then
        MsgBox "Indicar que el producto es de Cargo Automatico dentro del Combo"
        chkEliminaFijo.value = 0
        Me.chkFijo.SetFocus
    End If
   End If
End Sub

Private Sub txtAumPrec_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtAumPrec
End Sub

Private Sub chkFijo_Click()
    If chkFijo.value = 0 Then
        chkEliminaFijo.value = 0
    End If
End Sub

 

Private Sub txtEtiqueta_LostFocus()
   Call ValidaStr(txtEtiqueta)
End Sub
 
Private Sub txtQnCombo_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtQnCombo
End Sub




