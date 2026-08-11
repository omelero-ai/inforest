VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPedido 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8325
   ClientLeft      =   -660
   ClientTop       =   1710
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "frmPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Begin VB.PictureBox cmdOpcionPicture 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   10560
      Picture         =   "frmPedido.frx":030A
      ScaleHeight     =   615
      ScaleWidth      =   1170
      TabIndex        =   31
      Top             =   7560
      Width           =   1170
   End
   Begin VB.TextBox txtObservacion 
      Height          =   735
      Left            =   5310
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   5355
      Width           =   5145
   End
   Begin VB.TextBox txtPropiedad 
      Height          =   795
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   5355
      Width           =   5145
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Precios"
      Height          =   630
      Index           =   5
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Re Enviar Pedido"
      Height          =   630
      Index           =   6
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3345
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   9
      Left            =   10575
      Picture         =   "frmPedido.frx":6318
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6795
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   8
      Left            =   10575
      Picture         =   "frmPedido.frx":6BE2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6165
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Combos "
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   45
      TabIndex        =   15
      Top             =   6120
      Width           =   10455
      Begin VB.TextBox txtObservacionCombo 
         Height          =   690
         Left            =   5310
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1395
         Width           =   5055
      End
      Begin VB.TextBox txtPropiedadCombo 
         Height          =   690
         Left            =   5310
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   450
         Width           =   5055
      End
      Begin TrueOleDBGrid80.TDBGrid grdCombo_old 
         Height          =   870
         Left            =   6600
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1535
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
      Begin TrueOleDBGrid80.TDBGrid grdCombo 
         Height          =   1815
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   3201
         _LayoutType     =   4
         _RowHeight      =   26
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
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   7
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   15263202
         RowSubDividerColor=   16777215
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
         _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
         _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(20)  =   ":id=6,.fontname=Arial"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=34,.fontname=Arial"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
         _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
         _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=36,.fontname=Arial"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
         _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
         _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   5310
         TabIndex        =   20
         Top             =   1170
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Propiedad :"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   5310
         TabIndex        =   19
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Enviar Mensaje"
      Height          =   810
      Index           =   7
      Left            =   10575
      Picture         =   "frmPedido.frx":74AC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4050
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   2
      Left            =   10575
      Picture         =   "frmPedido.frx":7A36
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   675
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   1
      Left            =   10575
      Picture         =   "frmPedido.frx":8300
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   4
      Left            =   10575
      Picture         =   "frmPedido.frx":8BCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1935
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   10575
      Picture         =   "frmPedido.frx":9494
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1305
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Salir"
      Height          =   810
      Index           =   0
      Left            =   11040
      Picture         =   "frmPedido.frx":9D5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   1185
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla_old 
      Height          =   2145
      Left            =   6720
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3784
      _LayoutType     =   4
      _RowHeight      =   26
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
      Splits(0).ScrollBars=   3
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
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   4695
      Left            =   120
      TabIndex        =   33
      Top             =   360
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   8281
      _LayoutType     =   4
      _RowHeight      =   26
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
      Splits(0).ScrollBars=   3
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   16777215
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
      Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   7
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   16777215
      RowDividerColor =   15263202
      RowSubDividerColor=   16777215
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=Arial"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
      _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Arial"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
      _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(20)  =   ":id=6,.fontname=Arial"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
      _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
      _StyleDefs(43)  =   "Named:id=33:Normal"
      _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
      _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
      _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(52)  =   ":id=34,.fontname=Arial"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
      _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
      _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=36,.fontname=Arial"
      _StyleDefs(60)  =   "Named:id=37:Caption"
      _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(62)  =   "Named:id=38:HighlightRow"
      _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
      _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
      _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin VB.Label txtFecha 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   765
      TabIndex        =   26
      Top             =   0
      Width           =   2400
   End
   Begin VB.Label txtObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4620
      TabIndex        =   25
      Top             =   0
      Width           =   5865
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observación :"
      Height          =   240
      Left            =   3285
      TabIndex        =   24
      Top             =   22
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   45
      TabIndex        =   23
      Top             =   45
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observación :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   5355
      TabIndex        =   13
      Top             =   5085
      Width           =   1200
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Propiedad :"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   12
      Top             =   5085
      Width           =   990
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Niño(s) :"
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
      Index           =   1
      Left            =   10620
      TabIndex        =   10
      Top             =   5535
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adulto(s) :"
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
      Left            =   10620
      TabIndex        =   9
      Top             =   4905
      Width           =   705
   End
   Begin VB.Label txtAdulto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10620
      TabIndex        =   8
      Top             =   5175
      Width           =   945
   End
   Begin VB.Label txtNino 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10620
      TabIndex        =   7
      Top             =   5805
      Width           =   945
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      Height          =   195
      Index           =   5
      Left            =   8190
      TabIndex        =   1
      Top             =   5070
      Width           =   570
   End
   Begin VB.Label txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8835
      TabIndex        =   0
      Top             =   5025
      Width           =   1425
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset
Dim RsTitulo As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsCombo As Recordset
Dim RsComboPropiedad As Recordset
' Miguel Mendoza - 10062019
Dim iProductoVerificaComboMaximo As Integer
Dim iProductoVerificaComboMinimo As Integer
Dim iProductoVerificaComboUsado As Integer
Dim sProductoVerificaComboUsado As String
' *************************

Private Sub cmdSalir_Click()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Unload Me
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Dim nPos As Integer

    Select Case Index
           Case Is = 0  'Salir
                Unload Me
           Case Is = 1  'Primero
                MoverPuntero Primero, grdGrilla
                Asignar
           Case Is = 2  'Previo
                MoverPuntero previo, grdGrilla
                Asignar
           Case Is = 3  'Siguiente
                MoverPuntero siguiente, grdGrilla
                Asignar
           Case Is = 4  'Ultimo
                MoverPuntero Ultimo, grdGrilla
                Asignar
          
           Case Is = 5 ' Precios
                If RsGrilla.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If Supervisor("03") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                tUsuarioAutoriza = sVar1
                sCodigo = RsGrilla!tCodigoPedido
                sDescrip = RsGrilla!tItem
                frmPreciosDetalle.Show vbModal
                If wEnter Then
                   nPos = RsGrilla.AbsolutePosition
                   RsGrilla.Requery
                   RsGrilla.AbsolutePosition = nPos
                   txtTotal.Caption = Format(Calcular("SELECT sum(DPEDIDO.nVenta) AS Codigo From DPEDIDO WHERE DPEDIDO.tCodigoPedido='" & sPedido & "' AND DPEDIDO.tEstadoItem='N'", Cn), "#,###,##0.00")
                
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
                End If
                 
           Case Is = 6  'Re Imp. Pedido
                frmAreasImpresora.Show vbModal
                If Not wEnter Then
                   Exit Sub
                End If
                '************** Miguel Mendoza ************
                If Not VerificaComboMaximo Then Exit Sub
                '***************************************
                
                Dim nLP As Integer
                nLP = Calcular("select nReimpresion as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
                If nLimiteReimpresion <> 0 And nLimiteReimpresion <= nLP Then
                   If Supervisor("14") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
            
                Screen.MousePointer = vbHourglass
                Dim RsImpresion As Recordset
                Dim RsArea As Recordset
                Dim sSalon As String
                Isql = "select * from vPedido " & _
                       "Where Codigo = '" & sPedido & "' And lImprimeArea = 1 and (lImprime = 1 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 1 )) and tArea='" & sCodigo & "'" & _
                       "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo" 'tItemCombo"
'                       "ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                       
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                If RsImpresion.RecordCount = 0 Then
                   Screen.MousePointer = vbDefault
                   MsgBox "No hay nada por reimprimir", vbExclamation, sMensaje
                   Exit Sub
                End If
                Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "' and tArea='" & sCodigo & "'", Cn)
                sSalon = Calcular("select tSalon as Codigo from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
                If RsArea!tIcono = "" Or sSalon = RsArea!tIcono Or ((IsNull(sSalon) Or sSalon = "") And RsArea!nValor = 1) Then
                   If RsImpresion.RecordCount <> 0 Then
                      RsImpresion.MoveFirst
                      If sTipo = "CajaRapida" Then
                         ImprimePedido RsImpresion, "R", RsArea!timpresora, RsArea!Area, True, RsProductoPropiedad, RsComboPropiedad, "Rapido"
                      Else
                         ImprimePedido RsImpresion, "R", RsArea!timpresora, RsArea!Area, True, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                      End If
                    End If
                End If
                Cn.Execute "Update DPEDIDO Set nEnvio = nEnvio + 1 where tCodigoPedido = '" & sPedido & "' and lImprimeArea = 1 " 'and tArea='" & sCodigo & "'
                Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & sPedido & "'"
                
                Screen.MousePointer = vbDefault
                LimpiaRs
               
           Case Is = 7  'Mensaje
                frmMensaje.Caption = "Mensaje"
                frmMensaje.Show vbModal
                If Not wEnter Or sDescrip = "" Then
                   Exit Sub
                End If
                Dim RsTemp As Recordset
                frmPrecuentaImpresora.Show vbModal
                If Not wEnter Then
                   Exit Sub
                End If
                Isql = "SELECT dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.vMozo.Descripcion AS Mozo, dbo.vSalon.Descripcion + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tObservacion AS Observacion, dbo.MPEDIDO.nAdulto AS Pax " & _
                       "FROM dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO LEFT OUTER JOIN dbo.TMESA INNER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON " & _
                       "dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo where tCodigoPedido='" & sPedido & "'"
                Set RsTemp = Lib.OpenRecordset(Isql, Cn)
                
                
                'CESAR REGISTRAR MENSAJE
                Isql = "insert into TMENSAJEUSUARIO( " & _
                          "tCodigoPedido, tUsuario, tMensaje, tCaja, tMozo, tMesa, tTipoPedido, fRegistro) " & _
                          "values ('" & sPedido & "', " & _
                                 " '" & sUsuario & "', " & _
                                 " '" & sDescrip & "', " & _
                                 " '" & sCaja & "', " & _
                                 " '" & RsTemp!Mozo & "', " & _
                                 " '" & RsTemp!Mesa & "', " & _
                                 " '" & RsTemp!TipoPedido & "', " & _
                                        "getdate())"
            
                Cn.Execute Isql
                '----------------------------------
                                
                                
                ImprimeMensaje sCodigo, sDescrip, RsTemp

                Set RsTemp = Nothing
           Case Is = 8  'Anterior
                MoverPuntero previo, grdCombo
                'SubAsignar
                
           Case Is = 9  'Siguiente
                MoverPuntero siguiente, grdCombo
                'SubAsignar
               
    End Select
End Sub

Private Sub cmdOpcionPicture_Click(Index As Integer)
    cmdOpcion_Click Index
End Sub

Private Sub Form_Load()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If

    Screen.MousePointer = vbHourglass
    Centrar Me
   
    'Ingrese el Titulo
    Me.Caption = " Caja : " & sCaja & " - " & Trim(sUsuario) & " "
           
    'Ingrese el SubTitulo
    If sTipo = "CajaRapida" Then
       txtTotal.Caption = Format(Calcular("SELECT sum(nVenta) AS Codigo From " & frmCajaRapida.sDetalle & " WHERE tEstadoItem='N'", Cn), "#,###,##0.00")
    Else
       grdGrilla.Caption = " Pedido : " & sPedido
       txtTotal.Caption = Format(Calcular("SELECT sum(DPEDIDO.nVenta) AS Codigo From DPEDIDO WHERE DPEDIDO.tCodigoPedido='" & sPedido & "' AND DPEDIDO.tEstadoItem='N'", Cn), "#,###,##0.00")
    End If
    'Ingresar la Vista de la Grilla
    If sTipo = "CajaRapida" Then
       Isql = "SELECT " & frmCajaRapida.sComboDetalle & ".tCodigoPedido, " & frmCajaRapida.sComboDetalle & ".tProducto, " & frmCajaRapida.sComboDetalle & ".tItem, " & frmCajaRapida.sComboDetalle & ".tItemCombo, " & frmCajaRapida.sComboDetalle & ".tProductoCombo, " & frmCajaRapida.sComboDetalle & ".nCantidad, " & frmCajaRapida.sComboDetalle & ".tCodigoGrupo, " & frmCajaRapida.sComboDetalle & ".tCodigoSubGrupo, dbo.TPRODUCTO.tDetallado AS Producto, dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tCaja, " & frmCajaRapida.sComboDetalle & ".lImprimeArea, " & frmCajaRapida.sComboDetalle & ".lImprime, " & frmCajaRapida.sComboDetalle & ".nOrden, CONVERT(bit, ISNULL(DATALENGTH(" & frmCajaRapida.sComboDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, " & frmCajaRapida.sComboDetalle & ".tObservacion, " & frmCajaRapida.sComboDetalle & ".lCorte " & _
              "FROM " & frmCajaRapida.sComboDetalle & " LEFT OUTER JOIN (SELECT tItem, tItemCombo, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad From " & frmCajaRapida.sComboPropiedad & " " & _
              "GROUP BY tItem, tItemCombo) AS T1 ON " & frmCajaRapida.sComboDetalle & ".tItemCombo = T1.tItemCombo AND " & frmCajaRapida.sComboDetalle & ".tItem = T1.tItem LEFT OUTER JOIN dbo.TPRODUCTO ON " & frmCajaRapida.sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN dbo.MPEDIDO ON " & frmCajaRapida.sComboDetalle & ".tCodigoPedido = dbo.MPEDIDO.tCodigoPedido "
                            
       Set RsCombo = Lib.OpenRecordset(Isql, Cn)
       Set RsTitulo = Lib.OpenRecordset("select nAdulto, nNino from MPEDIDO where tCodigoPedido = ''", Cn)
       
       Isql = "SELECT " & frmCajaRapida.sDetalle & ".*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, CASE " & frmCajaRapida.sDetalle & ".nDescuento WHEN 0 THEN 0 ELSE " & frmCajaRapida.sDetalle & ".nDescuento * 100 / " & frmCajaRapida.sDetalle & ".nPrecioOficial END AS Descuento, " & _
              "dbo.TPRODUCTO.lDescuento AS lDescuento, dbo.TPRODUCTO.lModificable AS lModificable, CONVERT(bit, ISNULL(DATALENGTH(" & frmCajaRapida.sDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) As lPropiedad, " & frmCajaRapida.sDetalle & ".lCorte " & _
              "FROM " & frmCajaRapida.sDetalle & " LEFT OUTER JOIN (SELECT tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END as nPropiedad FROM " & frmCajaRapida.sProductoPropiedad & " GROUP by tItem ) as T1 " & _
              "ON " & frmCajaRapida.sDetalle & ".tItem = T1.tItem LEFT OUTER JOIN " & _
              "dbo.vCortesia ON " & frmCajaRapida.sDetalle & ".tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.TPRODUCTO ON " & frmCajaRapida.sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
              "ORDER BY " & frmCajaRapida.sDetalle & ".tCodigoPedido, " & frmCajaRapida.sDetalle & ".tItem"
       sPedido = frmCajaRapida.Pedido
    Else
       Set RsTitulo = Lib.OpenRecordset("select nAdulto, nNino, fregistro, tObservacion from MPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn)
       Set RsCombo = Lib.OpenRecordset("select * from vPedidoCombo where tCodigoPedido='" & sPedido & "'", Cn)
       Isql = "select * from vPedidoDetalle where tCodigoPedido = '" & sPedido & "' order by tItem"
    End If
    
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
        
    Call ConfGrilla(18, grdGrilla, "Itm", 2, "tItem", 430, 2, 0, "", _
                                   "Ord", 2, "nOrden", 400, 2, 0, "", _
                                   "-", 2, "lCorte", 250, 2, 4, "", _
                                   "Producto", 2, "Producto", 1900, 0, 0, "", _
                                   "P.Oficial", 2, "nPrecioOficial", 780, 1, 0, "###,###,##0.00", _
                                   "Dcto", 2, "nDescuento", 600, 1, 0, "##0.00", _
                                   "P.Venta", 2, "nPrecioVenta", 780, 1, 0, "###,###,##0.00", _
                                   "Cant.", 2, "nCantidad", 700, 1, 0, "##,##0.00", _
                                   "SubTotal", 2, "nVenta", 900, 1, 0, "###,###,##0.00", _
                                   "( % )", 2, "Descuento", 550, 1, 0, "##0.00", _
                                   "F", 2, "tFacturado", 270, 0, 0, "", _
                                   "E", 2, "lImprime", 270, 0, 4, "", _
                                   "P", 2, "lPropiedad", 270, 0, 4, "", _
                                   "O", 2, "lObservacion", 270, 0, 4, "", _
                                   "Documento", 2, "tDocumento", 1500, 0, 0, "@@@@@@@@@@@@@@@", _
                                   "Usuario", 2, "tUsuariod", 1500, 0, 0, "", _
                                   "Envio", 2, "fEnvio", 1600, 0, 0, "", _
                                   "Comanda", 2, "tComanda", 1500, 0, 0, "")
    Set grdGrilla.DataSource = RsGrilla
        
    Call ConfGrilla(7, grdCombo, "Itm", 2, "tItemCombo", 450, 2, 0, "", _
                                 "Ord", 2, "nOrden", 420, 2, 0, "", _
                                 "Producto", 2, "Producto", 2000, 0, 0, "", _
                                 "Cant.", 2, "nCantidad", 700, 1, 0, "##,##0.00", _
                                 "E", 2, "lImprime", 310, 0, 4, "", _
                                 "P", 2, "lPropiedad", 310, 0, 4, "", _
                                 "O", 2, "lObservacion", 310, 0, 4, "")
    Set grdCombo.DataSource = RsCombo
    
    If lAlmacen Then
       Dim xSql As String
       Dim RsOp As Recordset
       Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
       If RsOp.RecordCount > 0 Then
           'CASO CORP.FERNANDEZ
            If BDLink = "1" Then
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!Descripcion & "' as Operador, 0, " & " t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                       "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto =  t3x.tCodigoProducto " & _
                       "Where lNoDescargo = 1"
            Else
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!Descripcion & "' as Operador, 0, " & " t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                       "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto =  t3x.tCodigoProducto " & _
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
       
    'Producto vs Propiedad
    If sTipo = "CajaRapida" Then
       Isql = "SELECT " & frmCajaRapida.sProductoPropiedad & ".tItem, " & frmCajaRapida.sProductoPropiedad & ".tCodigoPropiedad, " & frmCajaRapida.sProductoPropiedad & ".tProducto, t1.Descripcion AS Descripcion, Operador , " & frmCajaRapida.sProductoPropiedad & ".ncantidad ncantidad " & _
              "FROM " & frmCajaRapida.sProductoPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & frmCajaRapida.sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo and " & frmCajaRapida.sProductoPropiedad & ".tProducto = T1.tProducto and " & frmCajaRapida.sProductoPropiedad & ".tenlace= t1.tenlace "
    Else
       'Producto Propiedad
       Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador , isnull(tproductopropiedad.ncantidad,1) ncantidad " & _
              "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
              "where tCodigoPedido = '" & sPedido & "'"
       txtAdulto.Caption = IIf(IsNull(RsTitulo!nAdulto), 0, RsTitulo!nAdulto)
       txtNino.Caption = IIf(IsNull(RsTitulo!nNino), 0, RsTitulo!nNino)
       txtFecha.Caption = IIf(IsNull(RsTitulo!fRegistro), "", RsTitulo!fRegistro)
       txtObs.Caption = IIf(IsNull(RsTitulo!tObservacion), "", RsTitulo!tObservacion)
       
    End If
    Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    'Combo Propiedad
    If sTipo = "CajaRapida" Then
       Isql = "SELECT " & frmCajaRapida.sComboPropiedad & ".tItem, " & frmCajaRapida.sComboPropiedad & ".tItemCombo, T1.Descripcion, T1.Operador ," & frmCajaRapida.sComboPropiedad & ".ncantidad  " & _
              "FROM " & frmCajaRapida.sComboPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & frmCajaRapida.sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & frmCajaRapida.sComboPropiedad & ".tProducto = T1.tProducto AND dbo." & frmCajaRapida.sComboPropiedad & ".tEnlace = T1.tEnlace"
    Else
       Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador , isnull(tcombopropiedad.ncantidad,1) ncantidad " & _
              "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
              "WHERE tcodigoPedido='" & sPedido & "'"
    End If
    Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    Asignar
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsGrilla = Nothing
    Set RsProductoPropiedad = Nothing
    Set RsTitulo = Nothing
    Set frmPedido = Nothing
End Sub

Public Sub Asignar()
   If Not RsGrilla.EOF Then
      RsProductoPropiedad.Filter = "tItem='" & RsGrilla!tItem & "'"
      txtPropiedad.Text = ""
      If Not RsProductoPropiedad.EOF Then
         RsProductoPropiedad.MoveFirst
         Do While Not RsProductoPropiedad.EOF
            txtPropiedad.Text = txtPropiedad.Text & IIf(RsProductoPropiedad!nCantidad = 1, "", "(" & RsProductoPropiedad!nCantidad & ")") & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ", "
            RsProductoPropiedad.MoveNext
         Loop
      End If
      txtObservacion.Text = IIf(IsNull(RsGrilla!tObservacion), "", RsGrilla!tObservacion)
      txtPropiedadCombo.Text = ""
      txtObservacionCombo.Text = ""
      RsCombo.Filter = "tItem='" & RsGrilla!tItem & "'"
   End If
End Sub

Public Sub SubAsignar()
   If Not RsCombo.EOF Then
      txtPropiedadCombo.Text = ""
      txtObservacionCombo.Text = ""
      If Not RsCombo.EOF And Not RsGrilla.EOF Then
         RsComboPropiedad.Filter = "tItem='" & RsGrilla!tItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
         If Not RsComboPropiedad.EOF Then
            RsComboPropiedad.MoveFirst
            Do While Not RsComboPropiedad.EOF
               txtPropiedadCombo.Text = txtPropiedadCombo.Text & IIf(RsComboPropiedad!nCantidad = 1, "", "(" & RsComboPropiedad!nCantidad & ")") & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
               RsComboPropiedad.MoveNext
            Loop
         End If
      End If
      txtObservacionCombo.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
   End If
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   SubAsignar
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Asignar
End Sub

' Miguel Mendoza - 10062019
Private Function VerificaComboMaximo() As Boolean
    Dim nPos As Integer
   VerificaComboMaximo = True
   With RsGrilla
      .Requery
      If .RecordCount > 0 Then
         .MoveFirst
         
         Do While Not .EOF
            iProductoVerificaComboMaximo = 0
            iProductoVerificaComboMinimo = 0
            iProductoVerificaComboUsado = 0

            iProductoVerificaComboMaximo = _
            Calcular("select isnull(ncombinacion,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

            iProductoVerificaComboMinimo = _
            Calcular("select isnull(ncombinacionminima,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

            If iProductoVerificaComboMaximo > 0 Then
               sProductoVerificaComboUsado = _
               Calcular("select isnull(tdetallado,'') codigo FROM tproducto WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

               iProductoVerificaComboUsado = _
               Calcular("select isnull(sum(nCantidad),0) codigo FROM CPEDIDO WHERE tCodigoPedido='" & sPedido & "' " & _
               "and tproducto='" & Trim(.Fields(3)) & "' and titem ='" & RsGrilla!tItem & "'", Cn)

               If iProductoVerificaComboUsado > iProductoVerificaComboMaximo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad máxima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMaximo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  nPos = RsGrilla.AbsolutePosition
                  RsGrilla.Requery
                  RsGrilla.AbsolutePosition = nPos
                  'grdGrilla.SetFocus
                  Exit Function
               End If

               If iProductoVerificaComboUsado < iProductoVerificaComboMinimo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad mínima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMinimo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  nPos = RsGrilla.AbsolutePosition
                  RsGrilla.Requery
                  RsGrilla.AbsolutePosition = nPos
                  'grdGrilla.SetFocus
                  Exit Function
               End If
            End If

            .MoveNext
         Loop
      End If
   End With
End Function
' *******************

