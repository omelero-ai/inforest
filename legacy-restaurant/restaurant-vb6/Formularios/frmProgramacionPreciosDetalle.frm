VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmProgramacionPreciosDetalle 
   Caption         =   "Programación Cambios de Precios Detalle"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   19185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalle 
      Height          =   8655
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   19095
      Begin VB.TextBox TxtValor2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12240
         TabIndex        =   47
         Text            =   "0"
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox TxtValor 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7680
         TabIndex        =   45
         Text            =   "0"
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton BtnRecargo 
         Caption         =   "Recargo"
         Height          =   495
         Left            =   8880
         TabIndex        =   44
         Top             =   970
         Width           =   1455
      End
      Begin VB.CommandButton BtnDescuento 
         Caption         =   "Descuento"
         Height          =   495
         Left            =   13440
         TabIndex        =   43
         Top             =   970
         Width           =   1455
      End
      Begin VB.CommandButton BtnRedondeo 
         Caption         =   "Redondear"
         Enabled         =   0   'False
         Height          =   495
         Left            =   17160
         TabIndex        =   41
         Top             =   970
         Width           =   1455
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
         Index           =   3
         Left            =   12000
         Picture         =   "frmProgramacionPreciosDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   160
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
         Left            =   10800
         Picture         =   "frmProgramacionPreciosDetalle.frx":00F2
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   160
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
         Left            =   9600
         Picture         =   "frmProgramacionPreciosDetalle.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   160
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
         Left            =   8400
         Picture         =   "frmProgramacionPreciosDetalle.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   160
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar Grupo"
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
         Height          =   615
         Index           =   4
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar Grupo"
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
         Height          =   615
         Index           =   7
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar SubGrupo"
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
         Height          =   615
         Index           =   5
         Left            =   1100
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar Producto"
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
         Height          =   615
         Index           =   6
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1440
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar SubGrupo"
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
         Height          =   615
         Index           =   8
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1440
         Width           =   950
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar Producto"
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
         Height          =   615
         Index           =   9
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   950
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
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   240
         Width           =   5655
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   665
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   315
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   54984707
         UpDown          =   -1  'True
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   810
         TabIndex        =   14
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   54984705
         CurrentDate     =   38587
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6375
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   18855
         _ExtentX        =   33258
         _ExtentY        =   11245
         _LayoutType     =   4
         _RowHeight      =   20
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(13)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Arial"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(42)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(43)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(47)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(48)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(49)  =   ":id=34,.fontname=Arial"
         _StyleDefs(50)  =   "Named:id=35:Footing"
         _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   "Named:id=36:Selected"
         _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(54)  =   ":id=36,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(55)  =   ":id=36,.fontname=Arial"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=825"
         _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=38,.fontname=Arial"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton BtnActual 
         Caption         =   "Precio Actual Registrado"
         Height          =   495
         Left            =   15720
         TabIndex        =   42
         Top             =   970
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Aplicar %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11040
         TabIndex        =   48
         Top             =   1065
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Aplicar %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   46
         Top             =   1065
         Width           =   1140
      End
      Begin VB.Label LblAplicado 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Por Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   16920
         TabIndex        =   39
         Top             =   195
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Aplicado:"
         Height          =   195
         Left            =   16200
         TabIndex        =   38
         Top             =   240
         Width           =   660
      End
      Begin VB.Label LblConfirmado 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "En Programación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   14040
         TabIndex        =   37
         Top             =   195
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   13440
         TabIndex        =   36
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblCanal5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   435
         Left            =   16245
         TabIndex        =   35
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label lblCanal4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   13770
         TabIndex        =   34
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label lblCanal3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   11295
         TabIndex        =   33
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label lblCanal2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
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
         Height          =   435
         Left            =   8835
         TabIndex        =   32
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label lblCanal1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
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
         Height          =   435
         Left            =   6360
         TabIndex        =   31
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   675
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   2400
         TabIndex        =   17
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   19125
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8760
      Width           =   19185
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Copiar"
         Height          =   615
         Index           =   11
         Left            =   7920
         TabIndex        =   49
         Top             =   80
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Confirmar"
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
         Left            =   6600
         Picture         =   "frmProgramacionPreciosDetalle.frx":0C58
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   80
         Width           =   1170
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Height          =   405
         Left            =   17280
         MaxLength       =   15
         TabIndex        =   30
         Text            =   " "
         Top             =   120
         Width           =   1440
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6285
         TabIndex        =   1
         Top             =   60
         Width           =   6345
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4860
            Picture         =   "frmProgramacionPreciosDetalle.frx":118A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5340
            Picture         =   "frmProgramacionPreciosDetalle.frx":16CC
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5820
            Picture         =   "frmProgramacionPreciosDetalle.frx":1C0E
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmProgramacionPreciosDetalle.frx":2150
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmProgramacionPreciosDetalle.frx":2692
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmProgramacionPreciosDetalle.frx":2BD4
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   8
            Top             =   180
            Width           =   3330
         End
      End
   End
End
Attribute VB_Name = "frmProgramacionPreciosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsPrecio As Recordset
Dim RsCanalesVenta As Recordset
Dim sTemporal As String
Dim nColumna As Integer
Dim i As Integer
Dim lPasa1 As Boolean

Private Sub BtnActual_Click()
Isql = "update TPROGRAMAPRECIOS_DETA set " & _
        "nPrecioVentaN1 = nPrecioVenta1, " & _
        "nPrecioVentaN2 = nPrecioVenta2, " & _
        "nPrecioVentaN3 = nPrecioVenta3, " & _
        "nPrecioVentaN4 = nPrecioVenta4, " & _
        "nPrecioVentaN5 = nPrecioVenta5 " & _
        " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql
  
Isql = "update TPROGRAMAPRECIOS_CAB set PorceAumento =0, PorceDescuento = 0 where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql

          Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
        Set grdGrilla.DataSource = RsPrecio
        
        MsgBox "Actualizado al Precio Actual Registrado", vbInformation, sMensaje
End Sub

Private Sub BtnDescuento_Click()
If (Trim(TxtValor2.Text) = "" Or Trim(TxtValor2.Text) = "0") Then
    TxtValor2.Text = "0"
    MsgBox "Ingrese un valor de descuento", vbExclamation, sMensaje
    Exit Sub
End If

Isql = "update TPROGRAMAPRECIOS_DETA set " & _
        "nPrecioVentaN1 = nPrecioVentaN1 - (( nPrecioVentaN1 * " & TxtValor2.Text & ") / 100), " & _
        "nPrecioVentaN2 = nPrecioVentaN2 - (( nPrecioVentaN2 * " & TxtValor2.Text & ") / 100), " & _
        "nPrecioVentaN3 = nPrecioVentaN3 - (( nPrecioVentaN3 * " & TxtValor2.Text & ") / 100), " & _
        "nPrecioVentaN4 = nPrecioVentaN4 - (( nPrecioVentaN4 * " & TxtValor2.Text & ") / 100), " & _
        "nPrecioVentaN5 = nPrecioVentaN5 - ((nPrecioVentaN5 * " & TxtValor2.Text & ") / 100) " & _
        " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql
  
  
Isql = "update TPROGRAMAPRECIOS_CAB set PorceDescuento = " & TxtValor2.Text & " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql
  
          Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
        Set grdGrilla.DataSource = RsPrecio
        
        MsgBox "Descuento realizado", vbInformation, sMensaje
        BtnRedondeo.Enabled = True
End Sub

Private Sub BtnRecargo_Click()
If (Trim(TxtValor.Text) = "" Or Trim(TxtValor.Text) = "0") Then
    TxtValor.Text = "0"
    MsgBox "Ingrese un valor de recargo", vbExclamation, sMensaje
    Exit Sub
End If

Isql = "update TPROGRAMAPRECIOS_DETA set " & _
        "nPrecioVentaN1 = nPrecioVentaN1 + (( nPrecioVentaN1 * " & TxtValor.Text & ") / 100), " & _
        "nPrecioVentaN2 = nPrecioVentaN2 + (( nPrecioVentaN2 * " & TxtValor.Text & ") / 100), " & _
        "nPrecioVentaN3 = nPrecioVentaN3 + (( nPrecioVentaN3 * " & TxtValor.Text & ") / 100), " & _
        "nPrecioVentaN4 = nPrecioVentaN4 + (( nPrecioVentaN4 * " & TxtValor.Text & ") / 100), " & _
        "nPrecioVentaN5 = nPrecioVentaN5 + ((nPrecioVentaN5 * " & TxtValor.Text & ") / 100) " & _
        " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql
  
  
Isql = "update TPROGRAMAPRECIOS_CAB set PorceAumento = " & TxtValor.Text & " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
Cn.Execute Isql
  
          Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
        Set grdGrilla.DataSource = RsPrecio
        
        MsgBox "Recargo realizado", vbInformation, sMensaje
        BtnRedondeo.Enabled = True
End Sub

Private Sub BtnRedondeo_Click()
If pais = "000" Or pais = "002" Then
    Isql = "update TPROGRAMAPRECIOS_DETA set " & _
            "nPrecioVentaN1 = round(convert(float,nPrecioVentaN1),0), " & _
            "nPrecioVentaN2 = round(convert(float,nPrecioVentaN2),0), " & _
            "nPrecioVentaN3 = round(convert(float,nPrecioVentaN3),0), " & _
            "nPrecioVentaN4 = round(convert(float,nPrecioVentaN4),0), " & _
            "nPrecioVentaN5 = round(convert(float,nPrecioVentaN5),0) " & _
            " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
    Cn.Execute Isql

Else
    Isql = "update TPROGRAMAPRECIOS_DETA set " & _
            "nPrecioVentaN1 = dbo.[ObteniendoRedondeo] (nPrecioVentaN1), " & _
            "nPrecioVentaN2 = dbo.[ObteniendoRedondeo] (nPrecioVentaN2), " & _
            "nPrecioVentaN3 = dbo.[ObteniendoRedondeo] (nPrecioVentaN3), " & _
            "nPrecioVentaN4 = dbo.[ObteniendoRedondeo] (nPrecioVentaN4), " & _
            "nPrecioVentaN5 = dbo.[ObteniendoRedondeo] (nPrecioVentaN5) " & _
            " where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
    Cn.Execute Isql

End If

  
          Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
        Set grdGrilla.DataSource = RsPrecio
        
        MsgBox "Redondeo realizado", vbInformation, sMensaje
        BtnRedondeo.Enabled = False
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmProgramacionPrecios.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmProgramacionPrecios.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmProgramacionPrecios.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmProgramacionPrecios.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmProgramacionPrecios.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmProgramacionPrecios.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmProgramacionPrecios.RsCabecera.AbsolutePosition & " de " & frmProgramacionPrecios.RsCabecera.RecordCount
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
    cmdOpcion(11).Enabled = False
End Sub
Private Sub cmdOpcion_Click(Index As Integer)
 Dim fechaservi As Date
 Dim fechaProg As Date
 Select Case Index
           Case Is = 0  'Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               grdGrilla.Enabled = True
               Iniciar
               txtNombre.SetFocus
                Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
                Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
                Set grdGrilla.DataSource = RsPrecio
                cmdOpcion(1).Enabled = True
           Case Is = 1  'Grabar
                Dim nCorrela As String
                
'               'Chequea Datos
               If txtNombre.Text = "" Then MsgBox "Ingrese el nombre de la Programación", vbExclamation, sMensaje: txtNombre.SetFocus: Exit Sub
               
               Dim RsBuscar As Recordset
               
               Isql = "SELECT count(*)Cantidad FROM  TPROGRAMAPRECIOS_CAB where fecha ='" & Format(dtpFecha.Value, "yyyy/MM/dd") & "' and hora = '" & Format(dtpHora.Value, "HH:mm") & "' AND tCodigoProgPrecio <> '" & txtCodigo.Text & "'"
                Set RsBuscar = Lib.OpenRecordset(Isql, Cn)
               
               If RsBuscar!Cantidad > 0 Then
                    MsgBox "No se puede Regisrar, ya existe una programación con la misma fecha y hora", vbExclamation, sMensaje
                    Exit Sub
               End If

               fechaservi = Calcular("select getdate() as Codigo from TPROGRAMAPRECIOS_CAB", Cn)
               
               If fechaservi > dtpFecha.Value & " " & dtpHora.Value Then
                    MsgBox "No se puede Regisrar, fecha y hora es menor a la fecha y hora actual del sistema", vbExclamation, sMensaje
                    Exit Sub
               End If
               
              If Sw Then
                 nCorrela = Calcular("select max(tCodigoProgPrecio) as Codigo from TPROGRAMAPRECIOS_CAB", Cn)
                 If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigo.Text = "00001"
                 Else
                     txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                 End If

                Cn.Execute "Insert into TPROGRAMAPRECIOS_CAB (tCodigoProgPrecio, tNomProgPrecio, Fecha, Hora, lActivo,lConfirmado,lAplicado, tUsuario, fRegistro) " & _
               " values ('" & txtCodigo.Text & "','" & txtNombre.Text & "','" & Format(dtpFecha.Value, "yyyy/MM/dd") & "','" & Format(dtpHora.Value, "HH:mm") & "','" & chkActivo.Value & "',0,0, '" & sUsuario & "', getdate())"
                
                cmdOpcion(4).Enabled = True
                cmdOpcion(5).Enabled = True
                cmdOpcion(6).Enabled = True
                cmdOpcion(7).Enabled = True
                cmdOpcion(8).Enabled = True
                cmdOpcion(9).Enabled = True
                cmdOpcion(10).Enabled = True
                cmdOpcion(11).Enabled = False
                
                Sw = False
                frmProgramacionPrecios.RsCabecera.Sort = "tCodigoProgPrecio ASC"
                frmProgramacionPrecios.RsCabecera.Requery
                frmProgramacionPrecios.RsCabecera.MoveLast
                
                MsgBox "Registro Guardado", vbInformation, sMensaje
                
                ActivarBotones (True)
                
                cmdTexto.Caption = "Registro " & IIf(frmProgramacionPrecios.RsCabecera.RecordCount = 0, 0, frmProgramacionPrecios.RsCabecera.AbsolutePosition) & " de " & frmProgramacionPrecios.RsCabecera.RecordCount
            Else
                Cn.Execute "Update TPROGRAMAPRECIOS_CAB  set tNomProgPrecio =  '" & txtNombre.Text & "' , Fecha = '" & Format(dtpFecha.Value, "yyyy/MM/dd") & "' , Hora = '" & Format(dtpHora.Value, "HH:mm") & "' where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
                
                nPos = frmProgramacionPrecios.RsCabecera.Bookmark
                frmProgramacionPrecios.RsCabecera.Requery
                
                If frmProgramacionPrecios.RsCabecera.RecordCount = 0 Then
                   frmProgramacionPrecios.RsCabecera.Filter = adFilterNone
                End If
                
                frmProgramacionPrecios.RsCabecera.Bookmark = nPos
                Screen.MousePointer = vbDefault
                
                MsgBox "Programación Modificada", vbInformation, sMensaje
               End If
               'Cn.Execute "update toferta set lnuevomodificado=1 where tOferta='" & txtCodigo & "'"
                         
                                
           Case Is = 2  'Eliminar
                If frmProgramacionPrecios.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
'
               If MsgBox("Seguro de Eliminar la Programación " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               Dim lAplica As Boolean
               lAplica = Calcular("select lAplicado as Codigo from TPROGRAMAPRECIOS_CAB", Cn)
               
               If lAplica = True Then
                    MsgBox "Registro no puede ser eliminado, ya fue aplicado esta programación", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'Cn.Execute "UPDATE TPROGRAMAPRECIOS_CAB set lActivo = 0, lconfirmado=0 where tCodigoProgPrecio = '" & txtCodigo & "'"
               
'              'Cambia el Delete
               Cn.Execute "delete from TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & txtCodigo & "'"
               Cn.Execute "delete from TPROGRAMAPRECIOS_CAB where tCodigoProgPrecio = '" & txtCodigo & "'"
                
               frmProgramacionPrecios.RsCabecera.Requery
               If frmProgramacionPrecios.RsCabecera.RecordCount <> 0 Then
                  frmProgramacionPrecios.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmProgramacionPrecios.RsCabecera.RecordCount = 0, 0, frmProgramacionPrecios.RsCabecera.AbsolutePosition) & " de " & frmProgramacionPrecios.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
                  MsgBox "Programación eliminada", vbInformation, sMensaje
           Case Is = 3  'Salir
                Unload Me
            Case Is = 4 ' Agregar Grupo
               If txtCodigo.Text = "" Then
                     MsgBox "Se debe generar el codigo, se debe grabar para poder continuar con la asignacion de grupos, subgrupos y productos", vbExclamation, sMensaje
                     Exit Sub
               End If

               Isql = "select Codigo, Descripcion from vGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tGrupo as Codigo FROM TPROGRAMAPRECIOS_DETA INNER JOIN dbo.TPRODUCTO ON TPROGRAMAPRECIOS_DETA.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoProgPrecio = '" & txtCodigo.Text & "' GROUP BY dbo.TPRODUCTO.tGrupo)"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Grupo", 2, "Descripcion", 6000, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "insert into TPROGRAMAPRECIOS_DETA (tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1, nPrecioVentaN1, nPrecioVenta2, nPrecioVentaN2, nPrecioVenta3, nPrecioVentaN3, nPrecioVenta4, nPrecioVentaN4, nPrecioVenta5, nPrecioVentaN5, tUsuario, fRegistro) " & _
               "select '" & txtCodigo.Text & "' , Codigo, nPrecioVenta as nPrecioVenta1, nPrecioVenta nPrecioVentaN1, nPrecioDelivery as nPrecioVenta2, nPrecioDelivery nPrecioVentaN2, nPrecioLlevar as nPrecioVenta3, nPrecioLlevar nPrecioVentaN3, nPrecioCanal4 as nPrecioVenta4, nPrecioCanal4 nPrecioVentaN4, " & _
               "nPrecioCanal5 as nPrecioVenta5, nPrecioCanal5 nPrecioVentaN5 , '" & sUsuario & "', getdate() from vProducto where lActivo=1 and tGrupo='" & sCodigo & "'"

               Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
               Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
            
               Set grdGrilla.DataSource = RsPrecio
               
               RsPrecio.Requery
               
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
               
          Case Is = 5 ' Agregar SubGrupo
          If txtCodigo.Text = "" Then
                     MsgBox "Se debe generar el codigo, se debe grabar para poder continuar con la asignacion de grupos, subgrupos y productos", vbExclamation, sMensaje
                     Exit Sub
               End If
               Isql = "select Codigo, Grupo, Descripcion from vSubGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tSubGrupo as Codigo FROM TPROGRAMAPRECIOS_DETA INNER JOIN dbo.TPRODUCTO ON TPROGRAMAPRECIOS_DETA.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoProgPrecio = '" & txtCodigo.Text & "' GROUP BY dbo.TPRODUCTO.tSubGrupo)"
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 3000, 0, 0, "", _
                                                      "SubGrupo", 2, "Descripcion", 3200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "insert into TPROGRAMAPRECIOS_DETA (tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1, nPrecioVentaN1, nPrecioVenta2, nPrecioVentaN2, nPrecioVenta3, nPrecioVentaN3, nPrecioVenta4, nPrecioVentaN4, nPrecioVenta5, nPrecioVentaN5, tUsuario, fRegistro) " & _
               "select '" & txtCodigo.Text & "' , Codigo, nPrecioVenta as nPrecioVenta1, nPrecioVenta nPrecioVentaN1, nPrecioDelivery as nPrecioVenta2, nPrecioDelivery nPrecioVentaN2, nPrecioLlevar as nPrecioVenta3, nPrecioLlevar nPrecioVentaN3, nPrecioCanal4 as nPrecioVenta4, nPrecioCanal4 nPrecioVentaN4, " & _
               "nPrecioCanal5 as nPrecioVenta5, nPrecioCanal5 nPrecioVentaN5, '" & sUsuario & "', getdate() from vProducto where lActivo=1 and tSubGrupo='" & sCodigo & "'"

                Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
               Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
            
               Set grdGrilla.DataSource = RsPrecio


               RsPrecio.Requery
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
                         
          Case Is = 6 ' Agregar Producto
          If txtCodigo.Text = "" Then
                     MsgBox "Se debe generar el codigo, se debe grabar para poder continuar con la asignacion de grupos, subgrupos y productos", vbExclamation, sMensaje
                     Exit Sub
               End If
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lActivo=1 and Codigo not in (select tCodigoProducto from TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & txtCodigo.Text & "' ) order by Descripcion"
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
               
              
               
               Cn.Execute "insert into TPROGRAMAPRECIOS_DETA (tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1, nPrecioVentaN1, nPrecioVenta2, nPrecioVentaN2, nPrecioVenta3, nPrecioVentaN3, nPrecioVenta4, nPrecioVentaN4, nPrecioVenta5, nPrecioVentaN5, tUsuario, fRegistro) " & _
               "select '" & txtCodigo.Text & "' , Codigo, nPrecioVenta as nPrecioVenta1, nPrecioVenta nPrecioVentaN1, nPrecioDelivery as nPrecioVenta2, nPrecioDelivery nPrecioVentaN2, nPrecioLlevar as nPrecioVenta3, nPrecioLlevar nPrecioVentaN3, nPrecioCanal4 as nPrecioVenta4, nPrecioCanal4 nPrecioVentaN4, " & _
                "nPrecioCanal5 as nPrecioVenta5, nPrecioCanal5 nPrecioVentaN5, '" & sUsuario & "', getdate() from vProducto where lActivo=1 and Codigo='" & sCodigo & "'"

                 Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
               Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
            
               Set grdGrilla.DataSource = RsPrecio
               
               RsPrecio.Requery
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
               
          Case Is = 7 ' Elimina Grupo
               Isql = "select Codigo, Descripcion from vGrupo where Codigo in (SELECT dbo.TPRODUCTO.tGrupo as Codigo FROM vTPROGRAMAPRECIOS_DETA INNER JOIN dbo.TPRODUCTO ON vTPROGRAMAPRECIOS_DETA.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoProgPrecio = '" & txtCodigo.Text & "' GROUP BY dbo.TPRODUCTO.tGrupo)"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Grupo", 2, "Descripcion", 6000, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "delete TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & txtCodigo.Text & "' and  tCodigoProducto in (SELECT dbo.TPRODUCTO.tCodigoProducto from TPRODUCTO WHERE tGrupo='" & sCodigo & "')"
               RsPrecio.Requery
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
          
          Case Is = 8 ' Elimina SubGrupo
               Isql = "select Codigo, Grupo, Descripcion from vSubGrupo where lActivo=1 and Codigo in (SELECT dbo.TPRODUCTO.tSubGrupo as Codigo FROM vTPROGRAMAPRECIOS_DETA INNER JOIN dbo.TPRODUCTO ON vTPROGRAMAPRECIOS_DETA.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoProgPrecio = '" & txtCodigo.Text & "' GROUP BY dbo.TPRODUCTO.tSubGrupo)"
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 3000, 0, 0, "", _
                                                      "SubGrupo", 2, "Descripcion", 3200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               
               Cn.Execute "delete TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & txtCodigo.Text & "' and tCodigoProducto in (SELECT dbo.TPRODUCTO.tCodigoProducto from TPRODUCTO WHERE tSubGrupo='" & sCodigo & "')"
               RsPrecio.Requery
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
               
          Case Is = 9 ' Elimina Producto
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where Codigo in (select tCodigoProducto from vTPROGRAMAPRECIOS_DETA  where tCodigoProgPrecio = '" & txtCodigo.Text & "') order by Descripcion"
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
               
               Cn.Execute "delete TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & txtCodigo.Text & "' and tCodigoProducto ='" & sCodigo & "'"
               RsPrecio.Requery
               txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
            Case Is = 10
                If (Trim(txtCodigo.Text) = "") Then
                    MsgBox "Guardar el registro antes de confirmar!", vbExclamation, sMensaje
                    Exit Sub
                End If
            
               fechaservi = Calcular("select TOP 1 getdate() as Codigo from TPROGRAMAPRECIOS_CAB", Cn)
               fechaProg = Calcular("select CONVERT(varchar, CONVERT(date, fecha),23) + ' ' + Hora + ':00' as Codigo from TPROGRAMAPRECIOS_CAB where tCodigoProgPrecio ='" & txtCodigo.Text & "'", Cn)
               
               If fechaservi > fechaProg Then
                    MsgBox "No se puede Regisrar, fecha y hora es menor a la fecha y hora actual del sistema", vbExclamation, sMensaje
                    Exit Sub
               End If
               
               
                If MsgBox("¿Desea Confirmar la Programación?" & Chr(13) & "despues de ser confirmado, no podra ser modificada", vbQuestion + vbYesNo) = vbYes Then
                    
                    Cn.Execute "UPDATE TPROGRAMAPRECIOS_CAB SET lConfirmado = 1, tUsuarioConfirma = '" & sUsuario & "', fRegistroConfirma = getdate() where tCodigoProgPrecio = '" & txtCodigo.Text & "'"
                    nPos = frmProgramacionPrecios.RsCabecera.Bookmark
                    frmProgramacionPrecios.RsCabecera.Requery
                    
                    If frmProgramacionPrecios.RsCabecera.RecordCount = 0 Then
                       frmProgramacionPrecios.RsCabecera.Filter = adFilterNone
                    End If
                    
                    frmProgramacionPrecios.RsCabecera.Bookmark = nPos
                    Screen.MousePointer = vbDefault
                    grdGrilla.Enabled = False
                    MsgBox "Programación Confirmada " & txtNombre.Text, vbInformation, sMensaje
                    Unload Me
                Else
                     Exit Sub
                End If
''           Se crea plantilla
            Case Is = 11   'Copiar   (13/05/2024)
                Sw = True
                Dim tCodigoProgPlantilla As String
'               variable de correlativo padre
                tCodigoProgPlantilla = txtCodigo.Text
                If Sw Then
                     nCorrela = Calcular("select max(tCodigoProgPrecio) as Codigo from TPROGRAMAPRECIOS_CAB", Cn)
                     If IsNull(nCorrela) Or nCorrela = "" Then
                         txtCodigo.Text = "00001"
                     Else
                         txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                     End If
                     
                     
'                   Inserta la cabecera de la plantilla
                    Cn.Execute "Insert into TPROGRAMAPRECIOS_CAB (tCodigoProgPrecio,tCodigoProgPlantilla, tNomProgPrecio, Fecha, Hora, lActivo,lConfirmado,lAplicado, tUsuario, fRegistro) " & _
                   " values ('" & txtCodigo.Text & "','" & tCodigoProgPlantilla & "','" & txtNombre.Text & "','" & Format(dtpFecha.Value, "yyyy/MM/dd") & "','" & Format(dtpHora.Value, "HH:mm") & "','" & chkActivo.Value & "',0,0, '" & sUsuario & "', getdate())"
                    
'                   Consulta para llamar los productos del correlativo padre
                    Isql = "select tCodigoProducto from TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio = '" & tCodigoProgPlantilla & "'"
                    Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
'                   Se itera e ingresa los productos al correlativo de la plantilla
                    Do While Not RsPrecio.EOF
                        sCodigo = RsPrecio.Fields("tCodigoProducto").Value
                        
                        ' realiza la inserción con el correlativo plantilla
                        Cn.Execute "insert into TPROGRAMAPRECIOS_DETA (tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1, nPrecioVentaN1, nPrecioVenta2, nPrecioVentaN2, nPrecioVenta3, nPrecioVentaN3, nPrecioVenta4, nPrecioVentaN4, nPrecioVenta5, nPrecioVentaN5, tUsuario, fRegistro) " & _
                        "select '" & txtCodigo.Text & "' , Codigo, nPrecioVenta as nPrecioVenta1, nPrecioVenta nPrecioVentaN1, nPrecioDelivery as nPrecioVenta2, nPrecioDelivery nPrecioVentaN2, nPrecioLlevar as nPrecioVenta3, nPrecioLlevar nPrecioVentaN3, nPrecioCanal4 as nPrecioVenta4, nPrecioCanal4 nPrecioVentaN4, " & _
                        "nPrecioCanal5 as nPrecioVenta5, nPrecioCanal5 nPrecioVentaN5, '" & sUsuario & "', getdate() from vProducto where lActivo=1 and Codigo='" & sCodigo & "'"
                    
                        RsPrecio.MoveNext
                    Loop
                    RsPrecio.Close

                    Sw = False
'                   Muestra  y refresaca la cabecera en la tabla
                    frmProgramacionPrecios.RsCabecera.Sort = "tCodigoProgPrecio ASC"
                    frmProgramacionPrecios.RsCabecera.Requery
                    frmProgramacionPrecios.RsCabecera.MoveLast
                    
                    
                    MsgBox "Registro Copiado", vbInformation, sMensaje
                End If
            
        End Select
End Sub
Public Sub Iniciar()
    chkActivo.Value = 1
    dtpHora.Value = "00:00"
    dtpFecha.Value = Date
    txtNombre.Text = ""
    txtCodigo.Text = ""
    
    txtNombre.Enabled = True
    dtpHora.Enabled = True
    dtpFecha.Enabled = True
    chkActivo.Enabled = True
    TxtValor.Enabled = True
    BtnRecargo.Enabled = True
    TxtValor2.Enabled = True
    BtnDescuento.Enabled = True
    BtnActual.Enabled = True
    BtnRedondeo.Enabled = True
    cmdOpcion(2).Enabled = True
'    Cn.Execute "delete from " & sTemporal
    RsPrecio.Requery
    txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
End Sub
Sub Asignar()
    With frmProgramacionPrecios.RsCabecera
        'Cuadro de Texto
        txtCodigo.Text = IIf(IsNull(!tCodigoProgPrecio), "", !tCodigoProgPrecio)
        dtpFecha.Value = IIf(IsNull(!Fecha), "", !Fecha)
        dtpHora.Value = IIf(IsNull(!Hora), "", !Hora)
        txtNombre.Text = IIf(IsNull(!tNomProgPrecio), "", !tNomProgPrecio)
        chkActivo.Value = IIf(IsNull(!lActivo), 0, IIf(!lActivo, 1, 0))
        TxtValor.Text = IIf(IsNull(!PorceAumento), "0", !PorceAumento)
        TxtValor2.Text = IIf(IsNull(!PorceDescuento), "0", !PorceDescuento)
        
        If (IIf(IsNull(!lConfirmado), 0, IIf(!lConfirmado, 1, 0)) = 0) Then
            LblConfirmado.Caption = "En Programación"
        Else
            LblConfirmado.Caption = "Confirmado"
        End If
        
        
        If (IIf(IsNull(!lAplicado), 0, IIf(!lAplicado, 1, 0)) = 0) Then
            LblAplicado.Caption = "Por Aplicar"
        Else
            LblAplicado.Caption = "Aplicado"
        End If
        
       
    If (!lConfirmado = 0) And (IIf(IsNull(!lAplicado), 0, IIf(!lAplicado, 1, 0)) = 0) Then
        
        cmdOpcion(4).Enabled = True
        cmdOpcion(5).Enabled = True
        cmdOpcion(6).Enabled = True
        cmdOpcion(7).Enabled = True
        cmdOpcion(8).Enabled = True
        cmdOpcion(9).Enabled = True
        grdGrilla.Enabled = True
        cmdOpcion(1).Enabled = True
        cmdOpcion(10).Enabled = True
        cmdOpcion(11).Enabled = False
        
        TxtValor.Enabled = True
        TxtValor2.Enabled = True
        BtnRecargo.Enabled = True
        BtnDescuento.Enabled = True
        BtnActual.Enabled = True
        BtnRedondeo.Enabled = True
    Else
        grdGrilla.Enabled = False
        cmdOpcion(1).Enabled = False
        cmdOpcion(4).Enabled = False
        cmdOpcion(5).Enabled = False
        cmdOpcion(6).Enabled = False
        cmdOpcion(7).Enabled = False
        cmdOpcion(8).Enabled = False
        cmdOpcion(9).Enabled = False
        cmdOpcion(10).Enabled = False
        cmdOpcion(11).Enabled = True
        
        TxtValor.Enabled = False
        TxtValor2.Enabled = False
        BtnRecargo.Enabled = False
        BtnDescuento.Enabled = False
        BtnActual.Enabled = False
        BtnRedondeo.Enabled = False
        dtpFecha.Enabled = False
        dtpHora.Enabled = False
        txtNombre.Enabled = False
        chkActivo.Enabled = False
        cmdOpcion(2).Enabled = False
    End If
    
    
'        Cn.Execute "delete from " & sTemporal
'        Cn.Execute "insert into " & sTemporal & " (tCodigoProducto) select tCodigoProducto from TPROGRAMAPRECIOS_DETA where tCodigoProgPrecio='" & txtCodigo.Text & "'"
        
        Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
        Set grdGrilla.DataSource = RsPrecio
     '   RsPrecio.Requery
       txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
        
       
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set RsGrupo = Nothing
    Set RsSubGrupo = Nothing
    Set RsPrecio = Nothing
    'Set RsFrecuencia = Nothing
   ' Set frmOfertaDetalle = Nothing
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    
    'Ingrese el SubTitulo
    Me.Caption = " Programación de Cambio de Precios "
    fraDetalle.Caption = Me.Caption
    
    'sTemporal = dbTemporal(sCaja, 1, "tCodigoProducto", "nVarChar(7)")
    
    'Canales de Venta
  'TCANALVENTA
  
   lblCanal1.Caption = "Canal - " & sBoton1 '= Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
   lblCanal2.Caption = "Canal - " & sBoton2 '= Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
   lblCanal3.Caption = "Canal - " & sBoton3 '= Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
   lblCanal4.Caption = "Canal - " & sBoton4 '= Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
   lblCanal5.Caption = "Canal - " & sBoton5 '= Calcular("select (case when lactivo = 1 then tDetallado else '' end) as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)
   
    
    If Sw = True Then
        Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
    Else
    
    
       Asignar
        Isql = "SELECT tCodigoProgPrecio, tCodigoProducto, Grupo, SubGrupo, Descripcion, nPrecioVenta , nPrecioVenta_Nuevo, nPrecioDelivery, nPrecioDelivery_Nuevo, nPrecioLlevar, nPrecioLlevar_Nuevo , nPrecioCanal4, nPrecioCanal4_Nuevo, nPrecioCanal5, nPrecioCanal5_Nuevo   FROM  vTPROGRAMAPRECIOS_DETA where tCodigoProgPrecio ='" & txtCodigo.Text & "'"
        Set RsPrecio = Lib.OpenRecordset(Isql, Cn)
    End If
    
    'Configuración de la Grilla
    Call ConfGrilla(14, grdGrilla, "Codigo", 2, "tCodigoProducto", 800, 2, 0, "", _
                                  "Grupo", 2, "Grupo", 1450, 0, 0, "", _
                                  "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "", _
                                  "Producto", 2, "Descripcion", 2200, 0, 0, "", _
                                  "Actual", 2, "nPrecioVenta", 1230, 1, 0, "###,##0.00", _
                                  "Nuevo", 2, "nPrecioVenta_Nuevo", 1230, 1, 0, "###,##0.00", _
                                  "Actual", 2, "nPrecioDelivery", 1230, 1, 0, "###,##0.00", _
                                  "Nuevo", 2, "nPrecioDelivery_Nuevo", 1230, 1, 0, "###,##0.00", _
                                  "Actual", 2, "nPrecioLlevar", 1230, 1, 0, "###,##0.00", _
                                  "Nuevo", 2, "nPrecioLlevar_Nuevo", 1230, 1, 0, "###,##0.00", _
                                  "Actual", 2, "nPrecioCanal4", 1230, 1, 0, "###,##0.00", _
                                  "Nuevo", 2, "nPrecioCanal4_Nuevo", 1230, 1, 0, "###,##0.00", _
                                  "Actual", 2, "nPrecioCanal5", 1230, 1, 0, "###,##0.00", _
                                  "Nuevo", 2, "nPrecioCanal5_Nuevo", 1230, 1, 0, "###,##0.00")


    Set grdGrilla.DataSource = RsPrecio
    
 
    
    grdGrilla.Columns(4).backColor = &HFFFF00
    grdGrilla.Columns(5).backColor = &HFFC0C0
    
    grdGrilla.Columns(6).backColor = &HC0C000
    grdGrilla.Columns(7).backColor = &HFFC0C0
    
    grdGrilla.Columns(8).backColor = &HC0E0FF
    grdGrilla.Columns(9).backColor = &HFFC0C0
    
    grdGrilla.Columns(10).backColor = &HC0FFC0
    grdGrilla.Columns(11).backColor = &HFFC0C0
    
    grdGrilla.Columns(12).backColor = &HFFFFFF
    grdGrilla.Columns(13).backColor = &HFFC0C0
    
     
    
     If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
     Else
       ActivarBotones (True)
       Asignar
    End If
      'lPasa1 = True
    txtTotal.Text = Format(RsPrecio.RecordCount, "###,##0")
    cmdTexto.Caption = "Registro " & frmProgramacionPrecios.RsCabecera.AbsolutePosition & " de " & frmProgramacionPrecios.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub grdGrilla_AfterColEdit(ByVal ColIndex As Integer)

    If Trim(grdGrilla.Columns(5)) = "" Or Not IsNumeric(Trim(grdGrilla.Columns(5))) Then
        grdGrilla.Columns(5) = Format("0", "0.000")
    End If

'    Validaciòn de columna 4 y 5 de precios, cambia al mismo valor todas las demas columnas de precios  13/05/2024
    If (ColIndex = 4 Or ColIndex = 4) Then

        grdGrilla.Columns(4) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(5) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(6) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(7) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(8) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(9) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(10) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(11) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(12) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")
        grdGrilla.Columns(13) = Format(grdGrilla.Columns(ColIndex).Value, "0.00")

    End If
    
'    RsCabecera!nAjuste = RsCabecera!nTotal - RsCabecera!nStockActual
'    Asignar
'    Agregar_Item
'    RsCabecera.Requery
'    RsCabecera.MoveFirst
'    RsCabecera.Find "tCodigoProducto='" & sCodigo & "'"
'    If Not RsCabecera.EOF Then
'        RsCabecera.MoveNext
'        Asignar
'    Else
'        RsCabecera.MoveFirst
'    End If
'    grdGrilla.col = 7
'    lPasa1 = True


End Sub
Private Sub grdGrilla_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    'lPasa1 = False
End Sub

Private Sub grdGrilla_Click()
    'Asignar
End Sub

Private Sub grdGrilla_ColEdit(ByVal ColIndex As Integer)

'   validaciòn para habilitar columnas 4 y 5 (13/05/2024)
    If (ColIndex <> 4 And ColIndex <> 5 And ColIndex <> 7 And ColIndex <> 9 And ColIndex <> 11 And ColIndex <> 13) Then
       ' cmdOpcion_Click 1
        SendKeys "{Esc}"
    End If
    
 
End Sub

Private Sub grdGrilla_DblClick()
    grdGrilla.AllowUpdate = True
    'Asignar
End Sub

