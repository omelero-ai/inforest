VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMantenimientoMesa 
   Caption         =   "Mantenimiento de Mesas"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   Icon            =   "frmMantenimientoMesa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOpciones 
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
      Height          =   1230
      Left            =   8820
      TabIndex        =   46
      Top             =   6975
      Width           =   3075
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
         Index           =   11
         Left            =   210
         Picture         =   "frmMantenimientoMesa.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   375
         Width           =   1260
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
         Index           =   12
         Left            =   1635
         Picture         =   "frmMantenimientoMesa.frx":21FC
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   375
         Width           =   1260
      End
   End
   Begin VB.Frame fraMesa 
      Caption         =   " Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   45
      TabIndex        =   44
      Top             =   0
      Width           =   5760
      Begin VB.CommandButton cmdMesa 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   510
         Width           =   915
      End
      Begin VB.Image imgFondo 
         Height          =   7815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5505
      End
   End
   Begin VB.Frame fraSalon 
      Caption         =   " Salones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2865
      Left            =   5850
      TabIndex        =   43
      Top             =   90
      Width           =   6075
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   4680
         Picture         =   "frmMantenimientoMesa.frx":22EE
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   225
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Modificar"
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
         Left            =   4680
         Picture         =   "frmMantenimientoMesa.frx":2820
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   900
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Left            =   4680
         Picture         =   "frmMantenimientoMesa.frx":2922
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1575
         Width           =   1260
      End
      Begin TrueOleDBGrid80.TDBGrid grdSalon 
         Height          =   2475
         Left            =   135
         TabIndex        =   50
         Top             =   270
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   4366
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
   End
   Begin VB.Frame fraMMesa 
      Caption         =   " Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   3900
      Left            =   5850
      TabIndex        =   27
      Top             =   3000
      Width           =   6045
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Mover"
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
         Index           =   8
         Left            =   4365
         Picture         =   "frmMantenimientoMesa.frx":2A24
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3150
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Left            =   2965
         Picture         =   "frmMantenimientoMesa.frx":2B26
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3150
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Modificar"
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
         Index           =   6
         Left            =   1565
         Picture         =   "frmMantenimientoMesa.frx":2C28
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3150
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Index           =   5
         Left            =   165
         Picture         =   "frmMantenimientoMesa.frx":2D2A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3150
         Width           =   1260
      End
      Begin TrueOleDBGrid80.TDBGrid grdMesa 
         Height          =   2700
         Left            =   135
         TabIndex        =   54
         Top             =   315
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   4763
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
   End
   Begin VB.Frame fraLeyenda 
      Caption         =   " Leyenda "
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
      Left            =   5850
      TabIndex        =   30
      Top             =   6975
      Width           =   2895
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Fuera Servicio :"
         Height          =   195
         Left            =   1350
         TabIndex        =   42
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label19 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   870
         Width           =   285
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Sucia :"
         Height          =   195
         Left            =   1965
         TabIndex        =   40
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Lista :"
         Height          =   195
         Left            =   540
         TabIndex        =   38
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   990
         TabIndex        =   37
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E2C28D&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   990
         TabIndex        =   36
         Top             =   555
         Width           =   285
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   555
         Width           =   285
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   990
         TabIndex        =   34
         Top             =   870
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ocupada :"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   585
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueada :"
         Height          =   195
         Left            =   1605
         TabIndex        =   32
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Reservada :"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         Top             =   900
         Width           =   870
      End
   End
   Begin VB.Frame fraSalonM 
      Caption         =   " Mantenimiento de Salones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2835
      Left            =   5850
      TabIndex        =   0
      Top             =   90
      Width           =   6075
      Begin VB.CheckBox chkPrecuenta 
         Alignment       =   1  'Right Justify
         Caption         =   "Precuenta Detallada :"
         Height          =   210
         Left            =   225
         TabIndex        =   6
         Top             =   1710
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
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
         Index           =   4
         Left            =   3240
         Picture         =   "frmMantenimientoMesa.frx":325C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1935
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Left            =   4590
         Picture         =   "frmMantenimientoMesa.frx":335E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1935
         Width           =   1260
      End
      Begin VB.TextBox txtSalon 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txtDetalladoL 
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
         Left            =   1980
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   990
         Width           =   2670
      End
      Begin VB.CheckBox chkActivoL 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   210
         Left            =   1215
         TabIndex        =   7
         Top             =   1980
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.TextBox txtResumidoL 
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
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   4
         Text            =   " "
         Top             =   1350
         Width           =   2670
      End
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   630
         Width           =   2670
         _ExtentX        =   4710
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Local :"
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   49
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   1170
         TabIndex        =   13
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   1350
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1080
         Width           =   1650
      End
   End
   Begin VB.Frame fraMMesaM 
      Caption         =   " Mantenimiento de Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   3900
      Left            =   5850
      TabIndex        =   14
      Top             =   3015
      Width           =   6045
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
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
         Index           =   9
         Left            =   3015
         Picture         =   "frmMantenimientoMesa.frx":3890
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2565
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Index           =   10
         Left            =   4410
         Picture         =   "frmMantenimientoMesa.frx":3992
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2565
         Width           =   1260
      End
      Begin VB.TextBox txtResumidoM 
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
         Left            =   1935
         MaxLength       =   15
         TabIndex        =   17
         Text            =   " "
         Top             =   1377
         Width           =   2595
      End
      Begin VB.TextBox txtDetalladoM 
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
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1008
         Width           =   2595
      End
      Begin VB.TextBox txtMesa 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txtAsiento 
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
         Left            =   1935
         MaxLength       =   15
         TabIndex        =   15
         Text            =   " "
         Top             =   639
         Width           =   1170
      End
      Begin VB.CheckBox chkActivoM 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   270
         Left            =   1185
         TabIndex        =   21
         Top             =   2100
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkFumador 
         Alignment       =   1  'Right Justify
         Caption         =   "Fumador :"
         Height          =   270
         Left            =   1020
         TabIndex        =   19
         Top             =   1746
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   1053
         Width           =   1650
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1422
         Width           =   1680
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   1215
         TabIndex        =   24
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Asientos :"
         Height          =   195
         Left            =   1110
         TabIndex        =   20
         Top             =   684
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmMantenimientoMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMesa As Recordset
Dim RsSalon As Recordset
Dim RsLocal As Recordset
Dim nIndexM As Integer
Dim nX As Integer
Dim nY As Integer
Dim wAgrega As Boolean

Private Sub cmdMesa_Click(Index As Integer)
     RsMesa.MoveFirst
     RsMesa.Move (Index - 1)
     txtMesa.Text = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
     nIndexM = Index
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Salon - Agregar
               wAgrega = True
               txtSalon.Text = ""
               txtDetalladoL.Text = ""
               txtResumidoL.Text = ""
               chkActivoL.value = 1
               fraSalon.Visible = False
                              
          Case Is = 1 ' Salon - Modificar
               If RsSalon.RecordCount = 0 Then
                  Exit Sub
               End If
               wAgrega = False
               
               txtSalon.Text = RsSalon!codigo
               txtDetalladoL.Text = RsSalon!Descripcion
               txtResumidoL.Text = RsSalon!tResumido
               cboLocal.BoundText = RsSalon!tLocal
               chkActivoL.value = IIf(RsSalon!lActivo, 1, 0)
               chkPrecuenta.value = IIf(RsSalon!nValor, 1, 0)
               
               fraSalon.Visible = False
               
          Case Is = 2 ' Salon - Eliminar
               'Cambia el MsgBox
               If RsSalon.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar el Salon " & RsSalon!Descripcion & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Verifica
               If RsMesa.RecordCount <> 0 Then
                  nPos = RsMesa.AbsolutePosition
                  RsMesa.MoveFirst
                  For i = 1 To RsMesa.RecordCount
                      If RsMesa!tEstadoMesa = "02" Or RsMesa!tEstadoMesa = "03" Then
                         MsgBox "Existen Mesas Asignadas o Reservadas", vbCritical, sMensaje
                         RsMesa.AbsolutePosition = nPos
                         Exit Sub
                      End If
                      RsMesa.MoveNext
                  Next i
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'SALON' and tCodigo = '" & RsSalon!codigo & "'"
               Cn.Execute "delete from TMESA where tSalon = '" & RsSalon!codigo & "'"
               RsSalon.Requery
               
               If RsSalon.RecordCount <> 0 Then
                  RsSalon.MoveLast
                  Desasigna RsMesa.RecordCount, cmdMesa
                  AsignaMesa
                  fraMesa.Caption = " " & RsSalon!Descripcion & " "
               Else
                  Desasigna RsMesa.RecordCount, cmdMesa
                  fraMesa.Caption = ""
               End If
               MsgBox "Registro Eliminado", vbExclamation, sMensaje
         
          Case Is = 3 ' Salon - Grabar
          
               'Chequea Datos
               If txtDetalladoL.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetalladoL.SetFocus: Exit Sub
               If txtResumidoL.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumidoL.SetFocus: Exit Sub
                    
               If wAgrega Then
                  'Asignar El Campo de Codificación
                   Set RsCodigo = Lib.OpenRecordset("select max(tCodigo) as Codigo from TTABLA where tTabla ='Salon' ", Cn)
                   If IsNull(RsCodigo!codigo) Then
                      txtSalon.Text = "01"
                   Else
                      txtSalon.Text = Lib.Correlativo(RsCodigo!codigo, 2)
                   End If
                   wAgrega = False
                   
                   'Cambiar el SQL
                   Isql = "insert into TTABLA( " & _
                          "tTabla, tCodigo, tDetallado, tResumido, tValor, nValor, lActivo) " & _
                          "values ('Salon', " & _
                                 " '" & txtSalon.Text & "', " & _
                                 " '" & txtDetalladoL.Text & "', " & _
                                 " '" & txtResumidoL.Text & "', " & _
                                 " '" & cboLocal.BoundText & "', " & _
                                 " '" & chkPrecuenta.value & "', " & _
                                        chkActivoL.value & ") "
            
                       Cn.Execute Isql
                       RsSalon.Requery
                       RsSalon.MoveLast
                       RsMesa.Requery
                       Desasigna RsMesa.RecordCount, cmdMesa
                       AsignaMesa
                       MsgBox "Registro Adicionado", vbInformation, sMensaje
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetalladoL.Text & "', " & _
                         "tResumido ='" & txtResumidoL.Text & "', " & _
                         "tValor='" & cboLocal.BoundText & "', " & _
                         "nValor =" & chkPrecuenta.value & ", " & _
                         "lActivo =" & chkActivoL.value & _
                         " where tTABLA = 'Salon' and tCodigo = '" & txtSalon.Text & "'"
                       
                   Cn.Execute Isql
                   pos = RsSalon.Bookmark
                   RsSalon.Requery
                   RsSalon.Bookmark = pos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
               fraSalon.Visible = True
                         
          Case Is = 4 ' Salon - Cancelar
               wAgrega = False
               fraSalon.Visible = True
          
          Case Is = 5 ' Mesa - Agregar
               wAgrega = True
               txtMesa.Text = ""
               txtDetalladoM.Text = ""
               txtResumidoM.Text = ""
               txtAsiento.Text = ""
               chkFumador.value = 1
               chkActivoL.value = 1
               fraMMesa.Visible = False
         
          Case Is = 6 ' Mesa - Modificar
               If RsMesa.RecordCount = 0 Then
                  Exit Sub
               End If
               wAgrega = False
               txtMesa.Text = RsMesa!tCodigoMesa
               txtDetalladoM.Text = RsMesa!tDetallado
               txtResumidoM.Text = RsMesa!tResumido
               txtAsiento.Text = RsMesa!nPersona
               chkFumador.value = IIf(RsMesa!lFumador = True, 1, 0)
               chkActivoM.value = IIf(RsMesa!lActivo = True, 1, 0)
               fraMMesa.Visible = False
                    
          Case Is = 7 ' Mesa - Eliminar
               'Cambia el MsgBox
               If RsMesa.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar la Mesa " & RsMesa!tResumido & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               Desasigna RsMesa.RecordCount, cmdMesa
               'Verifica
               If RsMesa!tEstadoMesa = "02" Or RsMesa!tEstadoMesa = "03" Then
                  MsgBox "Esta Mesa esta Asignada o Reservada", vbCritical, sMensaje
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TMESA where tCodigoMesa = '" & RsMesa!tCodigoMesa & "'"
               RsMesa.Requery
               AsignaMesa
               MsgBox "Registro Eliminado", vbExclamation, sMensaje
          
          Case Is = 8 ' Mesa - Mover
               If cmdOpcion(8).Caption = "Mover" Then
                  cmdMesa(nIndexM).DragMode = 1
                  cmdOpcion(8).Caption = "Actualizar"
                  For i = 1 To RsMesa.RecordCount
                      If i <> nIndexM Then
                         cmdMesa(i).Enabled = False
                      End If
                  Next i
               Else
                  cmdMesa(nIndexM).DragMode = 0
                  cmdOpcion(8).Caption = "Mover"
                  Isql = "update TMESA set " & _
                         "tX =" & nY & ", " & _
                         "ty =" & nX & " " & _
                         " where tCodigoMesa = '" & txtMesa.Text & "'"
                  Cn.Execute Isql
                  RsMesa.Requery
                  RsMesa.MoveFirst
                  For i = 1 To RsMesa.RecordCount
                      If i <> nIndexM Then
                              If RsMesa!tEstadoMesa = "02" Then
                                 cmdMesa(i).Enabled = False
                              Else
                                 cmdMesa(i).Enabled = True
                              End If
                      End If
                      RsMesa.MoveNext
                  Next i
                  RsMesa.MoveFirst
               End If
        
          Case Is = 9 ' Mesa - Cancelar
               wAgrega = False
               fraMMesa.Visible = True
          
          Case Is = 10 ' Mesa - Grabar
               'Chequea Datos
               If txtDetalladoM.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetalladoM.SetFocus: Exit Sub
               If txtResumidoM.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumidoM.SetFocus: Exit Sub
               If txtAsiento.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtAsiento.SetFocus: Exit Sub
                    
               If wAgrega Then
                  'Asignar El Campo de Codificación
                   Set RsCodigo = Lib.OpenRecordset("select max(tCodigoMesa) as Codigo from TMESA", Cn)
                   If IsNull(RsCodigo!codigo) Then
                      txtMesa.Text = "001"
                   Else
                      txtMesa.Text = Lib.Correlativo(RsCodigo!codigo, 3)
                   End If
                   wAgrega = False
                   
                   'Cambiar el SQL
                   Isql = "insert into TMESA( " & _
                          "tCodigoMesa, tDetallado, tResumido, tSalon, lFumador, tX, tY, nPersona, lActivo, tEstadoMesa, tUsuario, fRegistro) " & _
                          "values ('" & txtMesa.Text & "', " & _
                                 " '" & txtDetalladoM.Text & "', " & _
                                 " '" & txtResumidoM.Text & "', " & _
                                 " '" & RsSalon!codigo & "', " & _
                                        chkFumador.value & ", " & _
                                        300 & ", " & _
                                        120 & ", " & _
                                        val(txtAsiento.Text) & ", " & _
                                        chkActivoM.value & "," & _
                                 " '01', " & _
                                 " '" & sUsuario & "', " & _
                                 " getdate() )"
            
                   Cn.Execute Isql
                   RsMesa.Requery
                   RsMesa.MoveLast
                   txtMesa.Text = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
                   nIndexM = RsMesa.RecordCount
                   Desasigna RsMesa.RecordCount - 1, cmdMesa
                   AsignaMesa
                   MsgBox "Registro Adicionado", vbInformation, sMensaje
               Else
                  'Cambiar el SQL
                  Isql = "update TMESA set " & _
                         "tDetallado ='" & txtDetalladoM.Text & "', " & _
                         "tResumido ='" & txtResumidoM.Text & "', " & _
                         "lFumador =" & chkFumador.value & ", " & _
                         "nPersona =" & txtAsiento.Text & ", " & _
                         "lActivo =" & chkActivoM.value & ", lReplica=1 " & _
                         " where tCodigoMesa = '" & txtMesa.Text & "'"
                   Cn.Execute Isql
                   pos = RsMesa.Bookmark
                   RsMesa.Requery
                   RsMesa.Bookmark = pos
                   cmdMesa(nIndexM).Caption = RsMesa!tResumido
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               fraMMesa.Visible = True
                         
          Case Is = 12 ' salir
               Unload Me
   End Select
End Sub

Private Sub Form_Load()
   wAgrega = False
   Isql = "Select tCodigoMesa, tDetallado, TMESA.tResumido, tSalon, lFumador, tX, tY, nPersona, TMESA.lActivo, tUSuario, fRegistro, tEstadoMesa, vEstadoMesa.Descripcion as Estado " & _
          "FROM dbo.TMESA LEFT OUTER JOIN dbo.vEstadoMesa ON dbo.TMESA.tEstadoMesa = dbo.vEstadoMesa.Codigo where tCodigoMesa <> '000' order by tCodigoMesa"
   Set RsMesa = Lib.OpenRecordset(Isql, Cn)
   Set RsSalon = Lib.OpenRecordset("select * from vSalon", Cn)
   
   With cboLocal
        Set RsLocal = Lib.OpenRecordset("select * from vLocal", Cn)
        Set .RowSource = RsLocal
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
   End With
   
   Call ConfGrilla(2, grdSalon, "Local", 2, "Local", 1300, 0, 0, "", _
                                "Salon", 2, "Descripcion", 2500, 0, 0, "")
   Set grdSalon.DataSource = RsSalon
   
   Call ConfGrilla(5, grdMesa, "Mesa", 2, "tDetallado", 1600, 0, 0, "", _
                               "Pax", 2, "nPersona", 800, 1, 0, "", _
                               "Fumador", 2, "lFumador", 850, 2, 4, "", _
                               "Estado", 2, "Estado", 1100, 0, 0, "", _
                               "Activo", 2, "lActivo", 850, 2, 4, "")
   Set grdMesa.DataSource = RsMesa
   If RsSalon.RecordCount > 0 Then
      RsSalon.MoveFirst
      AsignaSalon
   End If
   AsignaMesa
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaMesa()
   cmdMesa(0).Visible = False
   
   If RsSalon.RecordCount = 0 Then
      Exit Sub
   End If
   
   RsMesa.Filter = "tSalon = '" & RsSalon!codigo & "'"
   If RsMesa.RecordCount = 0 Then
      Exit Sub
   End If
   
   RsMesa.MoveFirst
   cmdMesa(0).Visible = False
   
   For i = 1 To RsMesa.RecordCount
       'Cantidad de Mesas
       With cmdMesa(i)
            Load cmdMesa(i)
            .Top = RsMesa!tX
            .Left = RsMesa!tY
            .Visible = True
            cmdMesa(i).Caption = Trim(RsMesa!tResumido)
            
            Select Case RsMesa!tEstadoMesa
                   Case Is = "01" ' Libre
                        cmdMesa(i).Enabled = True
                        cmdMesa(i).backColor = vbLista
                   Case Is = "02" ' Ocupada
                        cmdMesa(i).Enabled = False
                        cmdMesa(i).backColor = vbOcupada
                   Case Is = "03" ' Reservada
                        cmdMesa(i).backColor = vbReservada
                        cmdMesa(i).Enabled = True
                   Case Is = "04" ' Sucia
                        cmdMesa(i).backColor = vbSucia
                        cmdMesa(i).Enabled = True
                   Case Is = "05" ' Bloqueada
                        cmdMesa(i).backColor = vbBloqueada
                        cmdMesa(i).Enabled = True
                   Case Is = "06" ' Fuera de Servicio
                        cmdMesa(i).backColor = vbFServicio
                        cmdMesa(i).Enabled = True
            End Select
            
       End With
       RsMesa.MoveNext
   Next i
   RsMesa.MoveFirst
End Sub

Public Sub Desasigna(Cantidad As Integer, cmdBoton As Object)
   For i = 1 To Cantidad
       Unload cmdBoton(i)
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMantenimientoMesa = Nothing
End Sub

Private Sub grdMesa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   nIndexM = RsMesa.AbsolutePosition
End Sub

Private Sub grdSalon_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignaSalon
   Desasigna RsMesa.RecordCount, cmdMesa
   AsignaMesa
End Sub

Private Sub imgFondo_DragDrop(Source As Control, X As Single, Y As Single)
  cmdMesa(nIndexM).Move X, Y
  nX = X
  nY = Y
End Sub

Public Sub AsignaSalon()
    If RsSalon.RecordCount > 0 Then
        fraMesa.Caption = " Mesas del Salon " & RsSalon!Descripcion & " "
        fraMMesa.Caption = " Mesas del Salon " & RsSalon!Descripcion & " "
        On Error Resume Next
        If Not IsNull(RsSalon!tIcono) Or RsSalon!tIcono <> "" Then
           imgFondo.Picture = LoadPicture(RsSalon!tIcono)
        Else
           imgFondo.Picture = LoadPicture()
        End If
    End If
End Sub
