VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmOfertaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8910
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11370
   Icon            =   "frmOfertaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11370
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
      Height          =   8160
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11300
      Begin VB.CheckBox chkRecomendacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Activar como recomendación:"
         Height          =   195
         Left            =   8400
         TabIndex        =   72
         Top             =   930
         Width           =   2415
      End
      Begin VB.TextBox txtUrlImagen 
         Height          =   285
         Left            =   4200
         TabIndex        =   71
         Top             =   210
         Width           =   6735
      End
      Begin VB.CommandButton btnStock 
         Caption         =   "Agregar Stock"
         Height          =   495
         Left            =   240
         TabIndex        =   69
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   8640
         MaxLength       =   24
         TabIndex        =   1
         Top             =   570
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   1845
         MaxLength       =   50
         TabIndex        =   0
         Text            =   " "
         Top             =   570
         Width           =   5655
      End
      Begin VB.CheckBox chkAcumulable 
         Alignment       =   1  'Right Justify
         Caption         =   "Acepta Descuentos :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3951
         Width           =   1800
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1185
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4254
         Width           =   855
      End
      Begin VB.CheckBox chkRatio 
         Alignment       =   1  'Right Justify
         Caption         =   "Ratio Oferta :"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   3345
         Width           =   1620
      End
      Begin VB.CheckBox chkExcluyente 
         Alignment       =   1  'Right Justify
         Caption         =   "Excluyente :"
         Height          =   195
         Left            =   870
         TabIndex        =   22
         Top             =   3648
         Width           =   1170
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   210
         Width           =   1170
      End
      Begin VB.CheckBox chkAutomatica 
         Alignment       =   1  'Right Justify
         Caption         =   "Oferta Automática :"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   930
         Width           =   1740
      End
      Begin VB.Frame Frame1 
         Caption         =   " Días de la Semana "
         Height          =   735
         Left            =   1860
         TabIndex        =   56
         Top             =   1230
         Width           =   9075
         Begin VB.CheckBox chkdia 
            Caption         =   "Lunes"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Martes"
            Height          =   195
            Index           =   2
            Left            =   1430
            TabIndex        =   4
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Miercoles"
            Height          =   195
            Index           =   3
            Left            =   2620
            TabIndex        =   5
            Top             =   300
            Width           =   1035
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Jueves"
            Height          =   195
            Index           =   4
            Left            =   3990
            TabIndex        =   6
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Viernes"
            Height          =   195
            Index           =   5
            Left            =   5180
            TabIndex        =   7
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Sábado"
            Height          =   195
            Index           =   6
            Left            =   6370
            TabIndex        =   8
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkdia 
            Caption         =   "Domingo"
            Height          =   195
            Index           =   7
            Left            =   7560
            TabIndex        =   9
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar Grupo"
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
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7440
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar Grupo"
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
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7440
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar SubGrupo"
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
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7440
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar Producto"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7440
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar SubGrupo"
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
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7440
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar Producto"
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
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7440
         Width           =   1170
      End
      Begin VB.TextBox txtMonto 
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
         Left            =   5737
         MaxLength       =   15
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   3300
         Width           =   1740
      End
      Begin VB.CheckBox chkMonto 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto Oferta :"
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   3345
         Width           =   1350
      End
      Begin VB.CheckBox chkdia 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Especial :"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   10
         Top             =   2085
         Width           =   1515
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Canal5"
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
         Index           =   5
         Left            =   9060
         TabIndex        =   29
         Top             =   4560
         Width           =   1725
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Canal4"
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
         Index           =   4
         Left            =   7260
         TabIndex        =   28
         Top             =   4560
         Width           =   1725
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   53
         Text            =   " "
         Top             =   7440
         Width           =   840
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Delivery"
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
         Index           =   2
         Left            =   3645
         TabIndex        =   26
         Top             =   4560
         Width           =   1725
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "En el Local"
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
         Index           =   1
         Left            =   1845
         TabIndex        =   25
         Top             =   4560
         Width           =   1725
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Para Llevar"
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
         Index           =   3
         Left            =   5445
         TabIndex        =   27
         Top             =   4560
         Width           =   1725
      End
      Begin VB.CheckBox chkPrecio 
         Alignment       =   1  'Right Justify
         Caption         =   "Precio Fijo :"
         Height          =   195
         Left            =   7815
         TabIndex        =   20
         Top             =   3345
         Width           =   1170
      End
      Begin VB.TextBox txtPrecio 
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
         Left            =   9195
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   3300
         Width           =   1740
      End
      Begin VB.CheckBox chkPermanente 
         Caption         =   "Oferta Permanente"
         Height          =   195
         Left            =   8820
         TabIndex        =   36
         Top             =   2935
         Width           =   1755
      End
      Begin VB.TextBox txtRatio 
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
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   3300
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker dtpFechaIni 
         Height          =   315
         Left            =   2835
         TabIndex        =   14
         Top             =   2875
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54263809
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   315
         Left            =   6435
         TabIndex        =   15
         Top             =   2875
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54263809
         CurrentDate     =   38587
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   2445
         Left            =   1860
         TabIndex        =   54
         Top             =   4920
         Width           =   9250
         _ExtentX        =   16325
         _ExtentY        =   4313
         _LayoutType     =   4
         _RowHeight      =   23
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   2835
         TabIndex        =   11
         Top             =   2025
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM"
         Format          =   54263811
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   315
         Left            =   6435
         TabIndex        =   13
         Top             =   2450
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm 'HRS'"
         Format          =   54263811
         UpDown          =   -1  'True
         CurrentDate     =   38587.9993055556
      End
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   315
         Left            =   2835
         TabIndex        =   12
         Top             =   2450
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm 'HRS'"
         Format          =   54263811
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin VB.Label Label3 
         Caption         =   "Url imagen:"
         Height          =   255
         Left            =   3240
         TabIndex        =   70
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Resumido :"
         Height          =   195
         Left            =   7680
         TabIndex        =   68
         Top             =   615
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "(Afecto a otro descuento)"
         Height          =   195
         Index           =   4
         Left            =   2220
         TabIndex        =   67
         Top             =   3960
         Width           =   6300
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta :"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   66
         Top             =   4560
         Width           =   1185
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "(Oferta que se carga en forma automática de acuerdo a los parámetros indicados)"
         Height          =   195
         Index           =   5
         Left            =   2220
         TabIndex        =   65
         Top             =   930
         Width           =   6300
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Franja Horaria :"
         Height          =   195
         Index           =   6
         Left            =   630
         TabIndex        =   63
         Top             =   2510
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Vigencia de la Oferta :"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   62
         Top             =   2935
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   1125
         TabIndex        =   61
         Top             =   255
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Frecuencia :"
         Height          =   195
         Index           =   19
         Left            =   825
         TabIndex        =   60
         Top             =   1245
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de la Oferta :"
         Height          =   195
         Left            =   195
         TabIndex        =   59
         Top             =   615
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora Final :"
         Height          =   195
         Index           =   24
         Left            =   5520
         TabIndex        =   58
         Top             =   2510
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora Inicial :"
         Height          =   195
         Index           =   23
         Left            =   1860
         TabIndex        =   57
         Top             =   2510
         Width           =   885
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "(No permite la venta de los productos, al finalizar la duracion de la oferta)"
         Height          =   195
         Index           =   7
         Left            =   2220
         TabIndex        =   55
         Top             =   3630
         Width           =   6300
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Index           =   1
         Left            =   1755
         TabIndex        =   52
         Top             =   2935
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Index           =   2
         Left            =   5430
         TabIndex        =   51
         Top             =   2935
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11310
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   8160
      Width           =   11370
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
         Index           =   3
         Left            =   9975
         Picture         =   "frmOfertaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   60
         Width           =   1170
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
         Left            =   8805
         Picture         =   "frmOfertaDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   60
         Width           =   1170
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
         Index           =   0
         Left            =   6465
         Picture         =   "frmOfertaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   60
         Width           =   1170
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
         Index           =   1
         Left            =   7635
         Picture         =   "frmOfertaDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6300
         TabIndex        =   42
         Top             =   60
         Width           =   6360
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmOfertaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmOfertaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmOfertaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5820
            Picture         =   "frmOfertaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5340
            Picture         =   "frmOfertaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4860
            Picture         =   "frmOfertaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
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
            Left            =   1440
            TabIndex        =   50
            Top             =   180
            Width           =   3330
         End
      End
   End
End
Attribute VB_Name = "frmOfertaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsProducto As Recordset
Dim RsFrecuencia As Recordset
Dim sTemporal As String
Dim nRatio As Double
Dim nMonto As Double
Dim nPrecio As Double
Dim sSemana As String
Dim nColumna As Integer
Dim i As Integer

Sub Asignar()
    With frmOferta.RsCabecera
    
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tOferta), "", !tOferta)
        txtNombre = IIf(IsNull(!tNombre), "", !tNombre)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtUrlImagen = IIf(IsNull(!UrlImagen), "", !UrlImagen)
        nRatio = IIf(IsNull(!nRatio), 0, !nRatio)
        nMonto = IIf(IsNull(!nMonto), 0, !nMonto)
        nPrecio = IIf(IsNull(!nPrecio), 0, !nPrecio)
                
        If nPrecio > 0 Then
           chkPrecio.Value = 1
           chkRatio.Value = 0
           chkMonto.Value = 0
           txtPrecio.Enabled = True
           txtRatio.Enabled = False
           txtMonto.Enabled = False
        ElseIf nMonto > 0 Then
           chkPrecio.Value = 0
           chkRatio.Value = 0
           chkMonto.Value = 1
           txtPrecio.Enabled = False
           txtRatio.Enabled = False
           txtMonto.Enabled = True
        Else
           chkPrecio.Value = 0
           chkRatio.Value = 1
           chkMonto.Value = 0
           txtPrecio.Enabled = False
           txtRatio.Enabled = True
           txtMonto.Enabled = False
        End If
        
        
        txtRatio = Format(nRatio, "##0.00")
        txtMonto = Format(nMonto, "###,##0.00")
        txtPrecio = Format(nPrecio, "###,##0.00")
        
       'Grilla
        Cn.Execute "delete from " & sTemporal
        Cn.Execute "insert into " & sTemporal & " (tCodigoProducto,Stock) select tCodigoProducto,Stock from TOFERTA where tOferta='" & txtCodigo & "'"
        RsProducto.Requery
        txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                            
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        'chkRecomendacion
        chkRecomendacion = IIf(!lRecomendacion = True, 1, 0)
        chkAcumulable = IIf(!lAcumulable = True, 1, 0)
        chkPermanente = IIf(!lPermanente = True, 1, 0)
        chkTipoPedido(1) = IIf(!lLocal = True, 1, 0)
        chkTipoPedido(2) = IIf(!lDelivery = True, 1, 0)
        chkTipoPedido(3) = IIf(!lLlevar = True, 1, 0)
        chkTipoPedido(4) = IIf(!lcanal4 = True, 1, 0)
        chkTipoPedido(5) = IIf(!lCanal5 = True, 1, 0)
        
        chkdia(0).Value = val(Mid(!tFrecuencia, 1, 1))
        chkdia(1).Value = val(Mid(!tFrecuencia, 2, 1))
        chkdia(2).Value = val(Mid(!tFrecuencia, 3, 1))
        chkdia(3).Value = val(Mid(!tFrecuencia, 4, 1))
        chkdia(4).Value = val(Mid(!tFrecuencia, 5, 1))
        chkdia(5).Value = val(Mid(!tFrecuencia, 6, 1))
        chkdia(6).Value = val(Mid(!tFrecuencia, 7, 1))
        chkdia(7).Value = val(Mid(!tFrecuencia, 8, 1))
        chkExcluyente.Value = IIf(!lExcluyente = True, 1, 0)
        chkAutomatica.Value = IIf(!lAutomatica = True, 1, 0)
        dtpFecha.Value = IIf(IsNull(!fFecha), Date, !fFecha)
        
        'Fechas
        dtpHoraIni.Value = !tHoraInicial
        dtpHoraFin.Value = !tHorafinal
                
        If chkPermanente.Value = 1 Then
           dtpFechaIni.Enabled = False
           dtpFechaFin.Enabled = False
           dtpFechaIni.Value = Date
           dtpFechaFin.Value = Date
        Else
           dtpFechaIni.Enabled = True
           dtpFechaFin.Enabled = True
           dtpFechaIni.Value = IIf(IsNull(!fFechaInicial), Date, !fFechaInicial)
           dtpFechaFin.Value = IIf(IsNull(!fFechafinal), Date, !fFechafinal)
        End If
        
        
    End With
End Sub

Private Sub btnStock_Click()

     Dim codigoProducto, CodOferta As String
     Dim Stock As Double
     
     nPos = RsProducto.AbsolutePosition
             
     codigoProducto = grdGrilla.Columns(0).Value
     CodOferta = txtCodigo.Text
         
     frmNumPad.Show vbModal
     'Stock = IIf(wEnter = True, Format(sDescrip, "###0.00"), "0.00")
     Stock = IIf(wEnter = True, sDescrip, "0.00")
     
     If Calcular("Select Count(*) AS CODIGO From TOFERTA where tCodigoProducto= '" & codigoProducto & "' and tOferta = '" & CodOferta & "'", Cn) = 0 Then
        MsgBox "El producto aún no ha sido registrado. Por favor, registre el producto antes de agregarle stock.", vbExclamation, sMensaje
        Exit Sub
     Else
         Cn.Execute "Update Toferta set Stock = " & Stock & " where tCodigoProducto= '" & codigoProducto & "' and tOferta = '" & CodOferta & "'"
         Asignar
     End If
     
     If nPos > 0 Then
        RsProducto.AbsolutePosition = nPos
     End If

End Sub

Private Sub chkdia_Click(Index As Integer)
   sSemana = ""
   If chkdia(0).Value = 1 Then
      Dim i As Integer
      For i = 1 To 7
          chkdia(i).Value = 0
          chkdia(i).Enabled = False
      Next i
      dtpFecha.Enabled = True
   Else
      For i = 1 To 7
          chkdia(i).Enabled = True
      Next i
      dtpFecha.Enabled = False
   End If
   sSemana = LTrim(str(chkdia(0).Value)) & LTrim(str(chkdia(1).Value)) & LTrim(str(chkdia(2).Value)) & LTrim(str(chkdia(3).Value)) & LTrim(str(chkdia(4).Value)) & LTrim(str(chkdia(5).Value)) & LTrim(str(chkdia(6).Value)) & LTrim(str(chkdia(7).Value))
End Sub

Private Sub chkMonto_Click()
   If chkMonto.Value = 1 Then
      chkRatio.Value = 0
      chkPrecio.Value = 0
      txtRatio.Enabled = False
      txtMonto.Enabled = True
      txtPrecio.Enabled = False
      nRatio = 0
      nPrecio = 0
   End If
   txtRatio = Format(nRatio, "##0.00")
   txtMonto = Format(nMonto, "###,##0.00")
   txtPrecio = Format(nPrecio, "###,##0.00")
End Sub

Private Sub chkPermanente_Click()
   If chkPermanente.Value = 0 Then
      dtpFechaIni.Enabled = True
      dtpFechaFin.Enabled = True
   Else
      dtpFechaIni.Enabled = False
      dtpFechaFin.Enabled = False
   End If
End Sub

Private Sub chkPrecio_Click()
   If chkPrecio.Value = 1 Then
      chkRatio.Value = 0
      chkMonto.Value = 0
      txtRatio.Enabled = False
      txtMonto.Enabled = False
      txtPrecio.Enabled = True
      nRatio = 0
      nMonto = 0
   End If
   txtRatio = Format(nRatio, "##0.00")
   txtMonto = Format(nMonto, "###,##0.00")
   txtPrecio = Format(nPrecio, "###,##0.00")
End Sub

Private Sub chkRatio_Click()
   If chkRatio.Value = 1 Then
      chkMonto.Value = 0
      chkPrecio.Value = 0
      txtRatio.Enabled = True
      txtMonto.Enabled = False
      txtPrecio.Enabled = False
      nMonto = 0
      nPrecio = 0
   End If
   txtRatio = Format(nRatio, "##0.00")
   txtMonto = Format(nMonto, "###,##0.00")
   txtPrecio = Format(nPrecio, "###,##0.00")
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmOferta.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmOferta.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmOferta.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmOferta.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmOferta.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmOferta.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmOferta.RsCabecera.AbsolutePosition & " de " & frmOferta.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Iniciar
               txtNombre.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtNombre.Text = "" Then MsgBox "Ingrese el nombre de la oferta", vbExclamation, sMensaje: txtNombre.SetFocus: Exit Sub
               
               If RsProducto.RecordCount = 0 Then MsgBox "Debe seleccionar un producto", vbExclamation, sMensaje: cmdOpcion(6).SetFocus: Exit Sub

               
               If Sw Then
                  If Not RsProducto.EOF Then
                     'Obtiene el Numero de Orden
                     nCorrela = Calcular("select max(tOferta) as Codigo from TOFERTA", Cn)
                     If IsNull(nCorrela) Or nCorrela = "" Then
                         txtCodigo.Text = "00001"
                     Else
                         txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                     End If
                     
                     Cn.Execute "insert into TOFERTA (tOferta, tCodigoProducto, tNombre, tResumido, tFrecuencia, fFecha, tHoraInicial, tHoraFinal, lAcumulable, nRatio, nMonto, nPrecio, lPermanente, fFechaInicial, fFechaFinal, lLocal, lDelivery, lLlevar, lCanal4, lCanal5, lexcluyente, lAutomatica, lActivo, tUsuario, fRegistro,Stock, UrlImagen, lRecomendacion) " & _
                                "select '" & txtCodigo.Text & "', tCodigoProducto, '" & txtNombre.Text & "', '" & txtResumido.Text & "', '" & sSemana & "', " & _
                                IIf(Mid(sSemana, 1, 1) = 1, "'" & Format(dtpFecha.Value, "yyyy/MM/dd") & "'", "Null") & ", '" & Format(dtpHoraIni.Value, "HH:mm") & "', '" & Format(dtpHoraFin.Value, "HH:mm") & "', " & _
                                chkAcumulable.Value & ", " & nRatio & ", " & nMonto & ", " & nPrecio & ", " & chkPermanente.Value & ", " & IIf(chkPermanente.Value, "Null", "'" & Format(dtpFechaIni.Value, "yyyy/MM/dd") & "'") & ", " & _
                                IIf(chkPermanente.Value, "Null", "'" & Format(dtpFechaFin.Value, "yyyy/MM/dd") & "'") & ", " & _
                                chkTipoPedido(1).Value & ", " & chkTipoPedido(2).Value & ", " & chkTipoPedido(3).Value & ", " & chkTipoPedido(4).Value & ", " & chkTipoPedido(5).Value & ", " & chkExcluyente.Value & ", " & chkAutomatica.Value & ", " & chkActivo.Value & ", '" & sUsuario & "', getdate(),Stock, '" & txtUrlImagen.Text & "', " & chkRecomendacion.Value & " from " & sTemporal
                     RsProducto.Requery
                     txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                  End If
                  Sw = False
                  frmOferta.RsCabecera.Sort = "tOferta ASC"
                  frmOferta.RsCabecera.Requery
                  frmOferta.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmOferta.RsCabecera.RecordCount = 0, 0, frmOferta.RsCabecera.AbsolutePosition) & " de " & frmOferta.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Cn.Execute "delete from TOFERTA where tOferta='" & txtCodigo & "'"
                  Cn.Execute "insert into TOFERTA (tOferta, tCodigoProducto, tNombre, tResumido, tFrecuencia, fFecha, tHoraInicial, tHoraFinal, lAcumulable, nRatio, nMonto, nPrecio, lPermanente, fFechaInicial, fFechaFinal, lLocal, lDelivery, lLlevar, lCanal4, lCanal5, lExcluyente, lAutomatica, lActivo, tUsuario, fRegistro,Stock, UrlImagen, lRecomendacion) " & _
                             "select '" & txtCodigo.Text & "', tCodigoProducto, '" & txtNombre.Text & "', '" & txtResumido.Text & "', '" & sSemana & "', " & _
                             IIf(Mid(sSemana, 1, 1) = 1, "'" & Format(dtpFecha.Value, "yyyy/MM/dd") & "'", "Null") & ", '" & Format(dtpHoraIni.Value, "HH:mm") & "', '" & Format(dtpHoraFin.Value, "HH:mm") & "', " & _
                             chkAcumulable.Value & ", " & nRatio & ", " & nMonto & ", " & nPrecio & ", " & chkPermanente.Value & ", " & IIf(chkPermanente.Value, "Null", "'" & Format(dtpFechaIni.Value, "yyyy/MM/dd") & "'") & ", " & _
                             IIf(chkPermanente.Value, "Null", "'" & Format(dtpFechaFin.Value, "yyyy/MM/dd") & "'") & ", " & _
                             chkTipoPedido(1).Value & ", " & chkTipoPedido(2).Value & ", " & chkTipoPedido(3).Value & ", " & chkTipoPedido(4).Value & ", " & chkTipoPedido(5).Value & ", " & chkExcluyente.Value & ", " & chkAutomatica.Value & ", " & chkActivo.Value & ", '" & sUsuario & "', getdate(),Stock, '" & txtUrlImagen.Text & "', " & chkRecomendacion.Value & " from " & sTemporal
                  RsProducto.Requery
                  txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                  
                  nPos = frmOferta.RsCabecera.Bookmark
                  frmOferta.RsCabecera.Requery
                  If frmOferta.RsCabecera.RecordCount = 0 Then
                     frmOferta.RsCabecera.Filter = adFilterNone
                  End If
                  frmOferta.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               Cn.Execute "update toferta set lnuevomodificado=1 where tOferta='" & txtCodigo & "'"
               
               btnStock.Enabled = True
                                        
          Case Is = 2 ' Eliminar
               If frmOferta.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar la Oferta " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TOferta where tOferta = '" & txtCodigo & "'"
               frmOferta.RsCabecera.Requery
               If frmOferta.RsCabecera.RecordCount <> 0 Then
                  frmOferta.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmOferta.RsCabecera.RecordCount = 0, 0, frmOferta.RsCabecera.AbsolutePosition) & " de " & frmOferta.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
                                             
          Case Is = 4 ' Agregar Grupo
               Isql = "select Codigo, Descripcion from vGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tGrupo)"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Grupo", 2, "Descripcion", 6000, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto,Stock) select tCodigoProducto,0 Stock from TPRODUCTO where lActivo=1 and tGrupo='" & sCodigo & "'"
                             
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
               
          Case Is = 5 ' Agregar SubGrupo
               Isql = "select Codigo, Grupo, Descripcion from vSubGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tSubGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tSubGrupo)"
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 3000, 0, 0, "", _
                                                      "SubGrupo", 2, "Descripcion", 3200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto,Stock) select tCodigoProducto,0 Stock from TPRODUCTO where lActivo=1 and tSubGrupo='" & sCodigo & "'"
                             
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
                         
          Case Is = 6 ' Agregar Producto
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lActivo=1 and Codigo not in (select tCodigoProducto from " & sTemporal & " ) order by Descripcion"
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
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto,Stock) values ('" & sCodigo & "',0)"
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
               
          Case Is = 7 ' Elimina Grupo
               Isql = "select Codigo, Descripcion from vGrupo where Codigo in (SELECT dbo.TPRODUCTO.tGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tGrupo)"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Grupo", 2, "Descripcion", 6000, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "delete " & sTemporal & " where tCodigoProducto in (SELECT dbo.TPRODUCTO.tCodigoProducto from TPRODUCTO WHERE tGrupo='" & sCodigo & "')"
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
          
          Case Is = 8 ' Elimina SubGrupo
               Isql = "select Codigo, Grupo, Descripcion from vSubGrupo where lActivo=1 and Codigo in (SELECT dbo.TPRODUCTO.tSubGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tSubGrupo)"
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 3000, 0, 0, "", _
                                                      "SubGrupo", 2, "Descripcion", 3200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "delete " & sTemporal & " where tCodigoProducto in (SELECT dbo.TPRODUCTO.tCodigoProducto from TPRODUCTO WHERE tSubGrupo='" & sCodigo & "')"
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
               
          Case Is = 9 ' Elimina Producto
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where Codigo in (select tCodigoProducto from " & sTemporal & " ) order by Descripcion"
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
               Cn.Execute "delete " & sTemporal & " where tCodigoProducto ='" & sCodigo & "'"
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
               btnStock.Enabled = False
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Ofertas "
    fraDetalle.Caption = Me.Caption
    
    sTemporal = dbTemporal(sCaja, 4, "tCodigoProducto", "nVarChar(7)", "Stock", "int", "UrlImagen", "nvarchar(max)", "lRecomendacion", "bit")
    Isql = "SELECT tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta, Stock, UrlImagen, lRecomendacion FROM " & sTemporal & " INNER JOIN dbo.vProducto ON tCodigoProducto = dbo.vProducto.Codigo"
    Set RsProducto = Lib.OpenRecordset(Isql, Cn)
            
    'Configuración de la Grilla
    Call ConfGrilla(6, grdGrilla, "Codigo", 2, "tCodigoProducto", 800, 2, 0, "", _
                                  "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                  "SubGrupo", 2, "SubGrupo", 2000, 0, 0, "", _
                                  "Producto", 2, "Descripcion", 2500, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                  "Stock", 2, "Stock", 700, 1, 0, "0")

    Set grdGrilla.DataSource = RsProducto
    chkTipoPedido(1).Caption = IIf(sBoton1 <> "", sBoton1, "n/d")
    chkTipoPedido(2).Caption = IIf(sBoton2 <> "", sBoton2, "n/d")
    chkTipoPedido(3).Caption = IIf(sBoton3 <> "", sBoton3, "n/d")
    chkTipoPedido(4).Caption = IIf(sBoton4 <> "", sBoton4, "n/d")
    chkTipoPedido(5).Caption = IIf(sBoton5 <> "", sBoton5, "n/d")
    chkTipoPedido(1).Enabled = IIf(sBoton1 <> "", True, False)
    chkTipoPedido(2).Enabled = IIf(sBoton2 <> "", True, False)
    chkTipoPedido(3).Enabled = IIf(sBoton3 <> "", True, False)
    chkTipoPedido(4).Enabled = IIf(sBoton4 <> "", True, False)
    chkTipoPedido(5).Enabled = IIf(sBoton5 <> "", True, False)
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
   Else
       ActivarBotones (True)
       Asignar
    End If
    
    txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    cmdTexto.Caption = "Registro " & frmOferta.RsCabecera.AbsolutePosition & " de " & frmOferta.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set RsGrupo = Nothing
    Set RsSubGrupo = Nothing
    Set RsProducto = Nothing
    Set RsFrecuencia = Nothing
    Set frmOfertaDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    RsProducto.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub txtMonto_LostFocus()
   nMonto = val(txtMonto.Text)
   txtMonto.Text = Format(nMonto, "###,##0.00")
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtMonto
End Sub


Private Sub txtPrecio_LostFocus()
   nPrecio = val(txtPrecio.Text)
   txtPrecio.Text = Format(nPrecio, "##0.00")
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtPrecio
End Sub

Private Sub txtRatio_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtRatio
End Sub

Private Sub txtRatio_LostFocus()

   nRatio = val(txtRatio.Text)
   
   If nRatio < 0 Or nRatio > 100 Then
        nRatio = 0
        Me.txtRatio.Text = "0"
        txtRatio.Text = Format(nRatio, "##0.00")
        MsgBox "Ratio no válido", vbInformation, sMensaje
        txtRatio.SetFocus
        Exit Sub
   End If
   
   txtRatio.Text = Format(nRatio, "##0.00")
End Sub

Public Sub Iniciar()
    chkActivo.Value = 1
    chkPermanente.Value = 1
    dtpHoraIni.Value = "00:00"
    dtpHoraFin.Value = "23:59"
    dtpFecha.Enabled = False
    sSemana = ""
    
    dtpFecha.Value = Date
    dtpFechaIni.Value = Date
    dtpFechaFin.Value = Date
    
    Cn.Execute "delete from " & sTemporal
    RsProducto.Requery
    txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    
    chkRatio.Value = 1
    chkMonto.Value = 0
    txtRatio.Enabled = True
    txtMonto.Enabled = False
    nRatio = 0
    nMonto = 0
    txtRatio = Format(nRatio, "##0.00")
    txtMonto = Format(nMonto, "##,###,##0.00")
    chkTipoPedido(1).Value = 1
    chkTipoPedido(2).Value = 1
    chkTipoPedido(3).Value = 1
    chkTipoPedido(4).Value = 1
    chkTipoPedido(5).Value = 1
    chkAutomatica.Value = 0
    chkActivo.Value = 1
    
End Sub
