VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPropiedadDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7320
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11505
   Icon            =   "frmPropiedadDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11505
   Begin VB.TextBox txtCodigoExterno 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   7680
      MaxLength       =   50
      TabIndex        =   57
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Visibilidad Canales de Venta"
      Height          =   2175
      Left            =   7680
      TabIndex        =   49
      Top             =   45
      Width           =   3735
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "Digital"
         Height          =   315
         Index           =   4
         Left            =   600
         TabIndex        =   55
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "Aplicativo"
         Height          =   315
         Index           =   3
         Left            =   600
         TabIndex        =   54
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "Recojo"
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   53
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "Delivery"
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   52
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "En el local"
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   51
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkVisibilidad 
         Caption         =   "Activar Visibilidad por Canal de Venta"
         Height          =   315
         Index           =   28
         Left            =   360
         TabIndex        =   50
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraInventario 
      Caption         =   " Afecta Inventario  "
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
      Left            =   7605
      TabIndex        =   31
      Top             =   4860
      Width           =   3840
      Begin VB.CommandButton cmdBusca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2925
         Picture         =   "frmPropiedadDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   322
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Insumo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   150
         Left            =   1935
         TabIndex        =   43
         Top             =   495
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Rec.Propiedad"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   150
         Left            =   1890
         TabIndex        =   42
         Top             =   315
         Width           =   960
      End
      Begin VB.Label txtDescripcionStock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   900
         TabIndex        =   41
         Top             =   757
         Width           =   2820
      End
      Begin VB.Label txtCodigoStock 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   900
         TabIndex        =   40
         Top             =   345
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " Código :"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Insumo :"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   817
         Width           =   735
      End
   End
   Begin VB.Frame fraValor 
      Caption         =   " Afecta Valor Venta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   7605
      TabIndex        =   29
      Top             =   3735
      Width           =   3840
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
         Height          =   330
         Left            =   990
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   405
         Width           =   1905
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "S/."
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
         Left            =   135
         TabIndex        =   46
         Top             =   450
         Width           =   735
      End
   End
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
      Height          =   6525
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   7530
      Begin VB.CheckBox chkEnvioTotem 
         Alignment       =   1  'Right Justify
         Caption         =   "Activa envio en Totem"
         Height          =   255
         Left            =   5025
         TabIndex        =   58
         Top             =   1860
         Width           =   2265
      End
      Begin VB.CheckBox chkPatioComida 
         Alignment       =   1  'Right Justify
         Caption         =   "Activa envio en Integración"
         Height          =   195
         Left            =   4680
         TabIndex        =   48
         Top             =   1620
         Width           =   2610
      End
      Begin VB.CheckBox chkCantidad 
         Alignment       =   1  'Right Justify
         Caption         =   "Solicita Cantidad al Seleccionar Propiedad:"
         Height          =   555
         Left            =   4680
         TabIndex        =   47
         Top             =   1080
         Width           =   2610
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
         Index           =   11
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5760
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
         Index           =   10
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5760
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
         Index           =   8
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5085
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
         Index           =   7
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5085
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
         Index           =   9
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5760
         Width           =   1170
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
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   27
         Text            =   " "
         Top             =   5085
         Width           =   840
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
         Index           =   6
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5085
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Marcar"
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
         Left            =   2520
         Picture         =   "frmPropiedadDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5085
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Desmarcar"
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
         Left            =   2520
         Picture         =   "frmPropiedadDetalle.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5760
         Width           =   1170
      End
      Begin VB.TextBox txtDetallado 
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   699
         Width           =   5070
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1068
         Width           =   3090
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo  :"
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   930
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   2835
         Left            =   90
         TabIndex        =   23
         Top             =   2160
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   5001
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
      Begin MSDataListLib.DataCombo cboOperador 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1437
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   1455
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Productos :"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   5130
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   738
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   1101
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   765
         TabIndex        =   19
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11445
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6570
      Width           =   11505
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
         Left            =   10245
         Picture         =   "frmPropiedadDetalle.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   9075
         Picture         =   "frmPropiedadDetalle.frx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   6735
         Picture         =   "frmPropiedadDetalle.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   7905
         Picture         =   "frmPropiedadDetalle.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6525
         TabIndex        =   11
         Top             =   60
         Width           =   6585
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmPropiedadDetalle.frx":13A0
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmPropiedadDetalle.frx":18E2
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmPropiedadDetalle.frx":1E24
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   6045
            Picture         =   "frmPropiedadDetalle.frx":2366
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5565
            Picture         =   "frmPropiedadDetalle.frx":28A8
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   5085
            Picture         =   "frmPropiedadDetalle.frx":2DEA
            Style           =   1  'Graphical
            TabIndex        =   12
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
            TabIndex        =   22
            Top             =   180
            Width           =   3555
         End
      End
   End
   Begin MSDataListLib.DataCombo cboArea 
      Height          =   315
      Left            =   8505
      TabIndex        =   44
      Top             =   6165
      Width           =   2820
      _ExtentX        =   4974
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
   Begin VB.Label Label8 
      Caption         =   "Codigo Externo"
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
      Left            =   7680
      TabIndex        =   56
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Area :"
      Height          =   195
      Left            =   7695
      TabIndex        =   45
      Top             =   6225
      Width           =   735
   End
End
Attribute VB_Name = "frmPropiedadDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsProducto As Recordset
Dim RsMensaje As Recordset
Dim RsMensajeVisibilidad As Recordset
Dim RsOperador As Recordset
Dim RsArea As Recordset
Dim sTemporal As String
Dim i As Integer

Sub LlenaCombos()
    With cboOperador
         Isql = "Select * from vOperador where isnull(lStockMenos,0)<>1 and lActivo = 1 order by descripcion"
         Set RsOperador = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOperador
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboArea
         Isql = "Select * from vArea where lActivo = 1 Union select '000', 'Sin Area','Sin Area','',1,'',0,0,0,0 order by codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
End Sub

Sub Asignar()

    
    With frmPropiedad.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        cboOperador.BoundText = IIf(IsNull(!tOperador), "", !tOperador)
                        
        'txtPrecio.Text = Format(Abs(IIf(IsNull(!nPrecio), 0, !nPrecio)), "###,###,##0.00")
        txtPrecio.Text = Format(IIf(IsNull(!nPrecio), 0, !nPrecio), "###,###,##0.00")
        txtCodigoStock.Caption = IIf(IsNull(!tEnlace), "", !tEnlace)
        txtDescripcionStock.Caption = ""
        Me.txtCodigoExterno.Text = IIf(IsNull(!tCodigoExterno), "", !tCodigoExterno)
        If IsNull(!tArea) Then
           cboArea.BoundText = "000"
        Else
           cboArea.BoundText = IIf(!tArea = "", "000", !tArea)
        End If
        
        If Len(txtCodigoStock.Caption) = 5 And lAlmacen = True Then
'           txtDescripcionStock.Caption = Calcular("Select tDescripcion as Codigo From vRecetaPropiedad Where lActivo = 1 and tLocal='" & sLocal & "' and tRecetaPropiedad='" & txtCodigoStock.Caption & "'", CnAlmacen)
           txtDescripcionStock.Caption = Calcular("Select tDescripcion as Codigo From vRecetaPropiedad Where tLocal='" & sLocal & "' and tRecetaPropiedad='" & txtCodigoStock.Caption & "'", CnAlmacen)
           txtDescripcionStock.ForeColor = &H800000
           txtCodigoStock.ForeColor = &H800000
           cboArea.Enabled = False
        ElseIf Len(txtCodigoStock.Caption) = 7 And lAlmacen = True Then
'           txtDescripcionStock.Caption = Calcular("Select tDetallado as Codigo From vProducto Where lDDirecto = 1 and tCodigoProducto ='" & txtCodigoStock.Caption & "'", CnAlmacen)
           txtDescripcionStock.Caption = Calcular("Select tDetallado as Codigo From vProducto Where tCodigoProducto ='" & txtCodigoStock.Caption & "'", CnAlmacen)
           txtDescripcionStock.ForeColor = &HC0&
           txtCodigoStock.ForeColor = &HC0&
           cboArea.Enabled = True
        Else
           cboArea.Enabled = False
        End If
                          
        'Grilla
        Cn.Execute "delete from " & sTemporal
        Cn.Execute "insert into " & sTemporal & " (tCodigoProducto, lEstado) select tProducto, isnull(lActivoProducto,1) from TPROPIEDAD where tCodigoPropiedad='" & txtCodigo & "'"
        RsProducto.Requery
        txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
        
  
        
'        Validad la exitencia de propiedad en table VISIBILIDADPROPIEDADXCANAL(31-05-2024)
        If Calcular("SELECT count(*) as codigo FROM VISIBILIDADPROPIEDADXCANAL WHERE tCodigoPropiedad = '" & txtCodigo.Text & "'", Cn) = 0 Then
                 
'           Insertar data en table VISIBILIDADPROPIEDADXCANAL(31-05-2024)-BR
            If (chkVisibilidad(28).Value = 1) Then
                Cn.Execute "insert into VISIBILIDADPROPIEDADXCANAL (tCodigoPropiedad, visibilidad, canal1, canal2, canal3, canal4, canal5) values ('" & txtCodigo & "',1, 0 ,0 ,0 ,0 ,0)"
                RsProducto.Requery

            Else
                Cn.Execute "insert into VISIBILIDADPROPIEDADXCANAL (tCodigoPropiedad, visibilidad, canal1, canal2, canal3, canal4, canal5) values ('" & txtCodigo & "',0, 0 ,0 ,0 ,0 ,0)"
                RsProducto.Requery
                
            End If
'           Obtener el valor guardado desde la tabla de propiedades -(31-05-2024 -BR)
            Isql = "select visibilidad from VISIBILIDADPROPIEDADXCANAL WHERE tCodigoPropiedad = '" & txtCodigo & "'"
            Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
        
            If (RsMensaje!Visibilidad) Then
                chkVisibilidad(28).Value = 1
            Else
                chkVisibilidad(28).Value = 0
            End If
        End If


    '    Obtener los nombres de los Canales de venta -(31-05-2024 -BR)
            Isql = "select tCodigoCanalVenta as Codigo, tDetallado as Descripcion, lActivo as Activo, lCanalCentralPedidos, lCanalDelivery, isnull(lMobileInforest,0) as lMobileInforest from TCANALVENTA"
            Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
            
            Isql = "select visibilidad,canal1, canal2, canal3, canal4, canal5 from VISIBILIDADPROPIEDADXCANAL where tCodigoPropiedad='" & txtCodigo & "'"
            Set RsMensajeVisibilidad = Lib.OpenRecordset(Isql, Cn)
    
            ' Verifica si se encontraron registros en el Recordset tabla tCodigoCanalVenta
            If Not RsMensaje.EOF Then
            Dim i As Integer
            i = 0 ' Inicializa el contador
            
           
            If (RsMensajeVisibilidad!Visibilidad = True) Then
             chkVisibilidad(28).Value = 1
            Else
             chkVisibilidad(28).Value = 0

                chkVisibilidad(0).Enabled = False
                     chkVisibilidad(0).Value = 0
                chkVisibilidad(1).Enabled = False
                     chkVisibilidad(1).Value = 0
                chkVisibilidad(2).Enabled = False
                     chkVisibilidad(2).Value = 0
                chkVisibilidad(3).Enabled = False
                     chkVisibilidad(3).Value = 0
                chkVisibilidad(4).Enabled = False
                     chkVisibilidad(4).Value = 0

            End If
             
            Dim t As Integer
            For t = 0 To 4
                If RsMensajeVisibilidad.Fields("canal" & (t + 1)).Value = True Then
                    chkVisibilidad(t).Value = 1
                Else
                    chkVisibilidad(t).Value = 0
                End If
            Next t

             
             
            Do While Not RsMensaje.EOF
                Dim codigo As String
                Dim Descripcion As String
                
                ' Asigna los valores de los campos a las variables
                codigo = RsMensaje.Fields("Codigo").Value
                Descripcion = RsMensaje.Fields("Descripcion").Value
                chkVisibilidad(i).Caption = Descripcion
'                If (RsMensajeVisibilidad.Fields(i).Value = True) Then
'                    chkVisibilidad(i).Value = 1
'                Else
'                    chkVisibilidad(i).Value = 0
'                End If
                
                i = i + 1
                RsMensaje.MoveNext
            Loop
    
            Else
                MsgBox "No se encontraron registros.", vbInformation, "Información de Canal de Venta"
            End If

        
'         Cierra el Recordset después de usarlo
        RsMensaje.Close
        Set RsMensaje = Nothing

        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)

         Me.chkPatioComida.Value = IIf(!lActivoPatio = True, 1, 0)
         
         Me.chkEnvioTotem.Value = IIf(!lActivoTotem = True, 1, 0)

        'Check Box
        Me.chkCantidad = IIf(!lSolicitacantidad = True, 1, 0)
    End With
End Sub

Private Sub cboOperador_Change()
   Dim tPrecio As String
   Dim lStock As Boolean
   
   RsOperador.MoveFirst
   RsOperador.Find "Codigo='" & cboOperador.BoundText & "'"
   If Not RsOperador.EOF Then
      lStock = IIf(IsNull(RsOperador!lStockMas), False, RsOperador!lStockMas)
      If IIf(IsNull(RsOperador!lValor), False, RsOperador!lValor) Then
         tPrecio = "+"
      Else
         tPrecio = ""
      End If
   Else
      tPrecio = ""
      lStock = False
   End If
   ActivaPrecio tPrecio
   
   If lStock And lAlmacen Then
      ActivaInventario True
      cboArea.BoundText = ""
   Else
      ActivaInventario False
      cboArea.BoundText = "000"
   End If
End Sub

Private Sub chkVisibilidad_Click(Index As Integer)
    Dim Isql As String
    Dim RsMensaje As Recordset
    
    If (chkVisibilidad(28).Value = 1) Then
        
     chkVisibilidad(0).Enabled = True
     chkVisibilidad(1).Enabled = True
     chkVisibilidad(2).Enabled = True
     chkVisibilidad(3).Enabled = True
     chkVisibilidad(4).Enabled = True
    Else
    
     chkVisibilidad(0).Enabled = False
          chkVisibilidad(0).Value = 0
     chkVisibilidad(1).Enabled = False
          chkVisibilidad(1).Value = 0
     chkVisibilidad(2).Enabled = False
          chkVisibilidad(2).Value = 0
     chkVisibilidad(3).Enabled = False
          chkVisibilidad(3).Value = 0
     chkVisibilidad(4).Enabled = False
          chkVisibilidad(4).Value = 0
    End If
    
   

End Sub

Private Sub cmdBusca_Click()
      Isql = "Select tCodigoProducto as Codigo, 'Insumo' as Tipo, tDetallado as Descripcion, UnidadEntrada as Dato From vPRODUCTO Where lDDirecto = 1 " & _
             "Union " & _
             "Select tRecetaPropiedad as Codigo, 'Receta' as Tipo, tDescripcion as Descripcion, Area as Dato From vRecetaPropiedad Where lActivo = 1 and tLocal='" & sLocal & "' order by Descripcion"
      
      frmBuscaAlmacen.nPredeterm = 2
      Call ConfGrilla(4, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 1000, 0, 0, "", _
                                                    "Tipo", 2, "Tipo", 1000, 0, 0, "", _
                                                    "Producto", 2, "Descripcion", 3400, 0, 0, "", _
                                                    "Unidad", 2, "Dato", 1800, 0, 0, "")
      frmBuscaAlmacen.Show vbModal
      If Not wEnter Then
         txtCodigoStock.Caption = ""
         txtDescripcionStock.Caption = ""
         cboArea.BoundText = ""
         Exit Sub
      End If
      txtCodigoStock.Caption = sCodigo
      txtDescripcionStock.Caption = sDescrip
      If Len(sCodigo) = 5 Then
         Dim sArea As String
         sArea = Calcular("select tCodigoArea as Codigo from MRECETAPROPIEDAD where tRecetaPropiedad='" & sCodigo & "' and tLocal='" & sLocal & "'", CnAlmacen)
         cboArea.BoundText = Calcular("select Codigo from varea where tValor='" & sArea & "'", Cn)
         If cboArea.BoundText = "" Then
            sCodigo = ""
            sDescrip = ""
            txtCodigoStock.Caption = sCodigo
            txtDescripcionStock.Caption = sDescrip
            MsgBox "Falta configurar el enlace de las áreas de producción con los Sub Almacenes", vbExclamation, sMensaje
            Exit Sub
         End If
         txtDescripcionStock.ForeColor = &H800000
         txtCodigoStock.ForeColor = &H800000
         cboArea.Enabled = False
      Else
         txtDescripcionStock.ForeColor = &HC0&
         txtCodigoStock.ForeColor = &HC0&
         cboArea.BoundText = ""
         cboArea.Enabled = True
      End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmPropiedad.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmPropiedad.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmPropiedad.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmPropiedad.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmPropiedad.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmPropiedad.grdGrilla
    End Select
    Asignar
    cmdTexto.Caption = "Registro " & frmPropiedad.RsCabecera.AbsolutePosition & " de " & frmPropiedad.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Dim RsMensaje As Recordset
   
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               Iniciar
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               
               'Validar y Guardar parametros de Configuraciòn de Visibilidad BR
               If (chkVisibilidad(28).Value = 1) Then
                 Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET visibilidad = 1  where tCodigoPropiedad='" & txtCodigo.Text & "'"
                 Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
                    ' Recorre todos los canales
                        For i = 0 To 4
                            ' Establece el nombre del campo del canal según el índice actual
                            Dim canalCampo As String
                            canalCampo = "canal" & (i + 1)
                            ' Verifica el valor del canal actual
                            Dim valorCanal As Integer
                            valorCanal = IIf(chkVisibilidad(i).Value = 1, 1, 0)
                            ' Construye y ejecuta la consulta SQL correspondiente para actualizar el canal actual
                            Select Case i
                                Case 0
                                    Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET " & canalCampo & " = " & valorCanal & " WHERE tCodigoPropiedad='" & txtCodigo.Text & "'"
                                Case 1
                                    Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET " & canalCampo & " = " & valorCanal & " WHERE tCodigoPropiedad='" & txtCodigo.Text & "'"
                                Case 2
                                    Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET " & canalCampo & " = " & valorCanal & " WHERE tCodigoPropiedad='" & txtCodigo.Text & "'"
                                Case 3
                                    Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET " & canalCampo & " = " & valorCanal & " WHERE tCodigoPropiedad='" & txtCodigo.Text & "'"
                                Case 4
                                    Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET " & canalCampo & " = " & valorCanal & " WHERE tCodigoPropiedad='" & txtCodigo.Text & "'"
                               
                            End Select
                            Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
                        Next i
   
               Else
                 Isql = "UPDATE VISIBILIDADPROPIEDADXCANAL SET visibilidad = 0,canal1 = 0,canal2 = 0 ,canal3 = 0 ,canal4 = 0,canal5 = 0    where tCodigoPropiedad='" & txtCodigo.Text & "'"
                 Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
               End If
            
               
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If cboOperador.Text = "" Then MsgBox "Ingrese el Operador de la Propiedad", vbExclamation, sMensaje: cboOperador.SetFocus: Exit Sub
               'If cboArea.Text = "" Then MsgBox "Ingrese el Area de la Propiedad", vbExclamation, sMensaje: cboOperador.SetFocus: Exit Sub
               
               'Cn.Execute "delete from " & sTemporal & " where lEstado=0"
               RsProducto.Requery
               
               If Sw Then
                  'Obtiene el Numero de Orden
                  Sw = False
                  If Not RsProducto.EOF Then
                     nCorrela = Calcular("select max(tCodigoPropiedad) as Codigo from TPROPIEDAD", Cn)
                     If IsNull(nCorrela) Or nCorrela = "" Then
                         txtCodigo.Text = "0001"
                     Else
                         txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                     End If
                     
                     Cn.Execute "insert into TPROPIEDAD (tCodigoPropiedad, tDetallado, tResumido, tOperador, tProducto, nPrecio, tEnlace, tArea, lActivo, tUsuario, fRegistro,lsolicitacantidad, lActivoProducto, lActivoPatio, tCodigoExterno, lActivoTotem) " & _
                                "select '" & txtCodigo.Text & "', '" & txtDetallado.Text & "', '" & txtResumido.Text & "', '" & cboOperador.BoundText & "', tCodigoProducto, " & IIf(fraValor.Enabled, CDbl(txtPrecio.Text), 0) & ", " & _
                                "'" & txtCodigoStock.Caption & "', '" & cboArea.BoundText & "', '" & chkActivo.Value & "', '" & sUsuario & "', getdate(), '" & chkCantidad.Value & "', lestado, " & chkPatioComida.Value & ", '" & Trim(Me.txtCodigoExterno.Text) & "', " & chkEnvioTotem.Value & " from " & sTemporal & " "
                     RsProducto.Requery
                     txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                      Else
                                    MsgBox "Ingrese por lo menos producto", vbExclamation, sMensaje
                                    Sw = True
                                    Exit Sub
                     End If
 
      
                  
                  frmPropiedad.RsCabecera.Sort = "Codigo ASC"
                  frmPropiedad.RsCabecera.Requery
                  frmPropiedad.RsCabecera.MoveLast
                  'frmPropiedad.RsCabecera.
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmPropiedad.RsCabecera.RecordCount = 0, 0, frmPropiedad.RsCabecera.AbsolutePosition) & " de " & frmPropiedad.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Cn.Execute "delete from TPROPIEDAD where tCodigoPropiedad='" & txtCodigo & "'"
                  Cn.Execute "insert into TPROPIEDAD (tCodigoPropiedad, tDetallado, tResumido, tOperador, tProducto, nPrecio, tEnlace, tArea, lActivo, tUsuario, fRegistro, lReplica,lsolicitacantidad,lActivoProducto, lActivoPatio, tCodigoExterno, lActivoTotem) " & _
                             "select '" & txtCodigo.Text & "', '" & txtDetallado.Text & "', '" & txtResumido.Text & "', '" & cboOperador.BoundText & "', tCodigoProducto, " & IIf(fraValor.Enabled, CDbl(txtPrecio.Text), 0) & ", " & _
                             "'" & txtCodigoStock.Caption & "', '" & cboArea.BoundText & "', '" & chkActivo.Value & "', '" & sUsuario & "', getdate(), 1, '" & chkCantidad.Value & "',lestado, " & chkPatioComida.Value & ", '" & Trim(Me.txtCodigoExterno.Text) & "', " & chkEnvioTotem.Value & " from " & sTemporal & " " '-- where lEstado=1
                  
                  RsProducto.Requery
                  txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                                                                        
                  nPos = frmPropiedad.RsCabecera.Bookmark
                  If frmPropiedad.RsCabecera.RecordCount = 0 Then
                     frmPropiedad.RsCabecera.Filter = adFilterNone
                  End If
                   
                   frmPropiedad.RsCabecera.Requery
                   frmPropiedad.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                Cn.Execute "update TPROPIEDAD set lNuevoModificado=1 where tCodigoPropiedad='" & txtCodigo & "'"
            
          Case Is = 2 ' Eliminar
               If frmPropiedad.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Propiedad " & txtDetallado & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If frmPropiedad.RsCabecera.Bookmark > 1 Then
                   nPos = frmPropiedad.RsCabecera.Bookmark - 1
                Else
                   nPos = 0
               End If
               'Cambia el Delete
               Cn.Execute "delete from TPROPIEDAD where tCodigopropiedad = '" & txtCodigo & "'"
               frmPropiedad.RsCabecera.Requery
               
               If nPos <> 0 Then
                frmPropiedad.RsCabecera.Bookmark = nPos
               End If
               
               If frmPropiedad.RsCabecera.RecordCount <> 0 Then
                  'frmPropiedad.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmPropiedad.RsCabecera.RecordCount = 0, 0, frmPropiedad.RsCabecera.AbsolutePosition) & " de " & frmPropiedad.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
          
          Case Is = 4 ' Desmarcar
               Cn.Execute "update " & sTemporal & " set lEstado=0"
               RsProducto.Requery
   
          Case Is = 5 ' Marcar
               Cn.Execute "update " & sTemporal & " set lEstado=1"
               RsProducto.Requery
   
          Case Is = 6 ' Agregar Grupo
               Isql = "select Codigo, Descripcion from vGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tGrupo)"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Grupo", 2, "Descripcion", 6000, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto, lEstado) select tCodigoProducto, 1 as lEstado from TPRODUCTO where lActivo=1 and tGrupo='" & sCodigo & "'"
                             
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
          Case Is = 7 ' Agregar SubGrupo
               Isql = "select Codigo, Grupo, Descripcion from vSubGrupo where lActivo=1 and Codigo not in (SELECT dbo.TPRODUCTO.tSubGrupo as Codigo FROM " & sTemporal & " INNER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto GROUP BY dbo.TPRODUCTO.tSubGrupo)"
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 3000, 0, 0, "", _
                                                      "SubGrupo", 2, "Descripcion", 3200, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto, lEstado) select tCodigoProducto, 1 as lEstado from TPRODUCTO where lActivo=1 and tSubGrupo='" & sCodigo & "'"
                             
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
          
          Case Is = 8 ' Agregar Producto
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
               Cn.Execute "insert into " & sTemporal & " (tCodigoProducto, lEstado) values ('" & sCodigo & "',1)"
               RsProducto.Requery
               txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
               
          Case Is = 9 ' Elimina Grupo
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
          
          Case Is = 10 ' Elimina SubGrupo
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
               
          Case Is = 11 ' Elimina Producto
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
               
   End Select

End Sub

Private Sub Form_Load()
    Dim RsMensaje As Recordset
    Dim RsMensajeVisibilidad As Recordset
    Dim activaVisibilidad As Boolean

    
    Screen.MousePointer = vbHourglass
    Centrar Me
    
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Propiedades "
    fraDetalle.Caption = Me.Caption
    
    sTemporal = dbTemporal(sCaja, 2, "tCodigoProducto", "nVarChar(7)", "lEstado", "bit")
    Isql = "SELECT tCodigoProducto, Grupo, SubGrupo, Descripcion, lEstado FROM " & sTemporal & " INNER JOIN dbo.vProducto ON tCodigoProducto = dbo.vProducto.Codigo"
    Set RsProducto = Lib.OpenRecordset(Isql, Cn)
            
    'Configuración de la Grilla
    Call ConfGrilla(5, grdGrilla, "Codigo", 2, "tCodigoProducto", 700, 2, 0, "", _
                                  "Grupo", 2, "Grupo", 1500, 0, 0, "", _
                                  "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "", _
                                  "Producto", 2, "Descripcion", 2300, 0, 0, "", _
                                  "Estado", 2, "lEstado", 700, 2, 4, "")
    Set grdGrilla.DataSource = RsProducto
    
    'Llena todos los Combos
    LlenaCombos
    
    ActivaPrecio ""
    cboArea.Enabled = False
    ActivaInventario False
    lblMoneda.Caption = "en " & sMonN
    If Sw = True Then
       Iniciar
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    cmdTexto.Caption = "Registro " & frmPropiedad.RsCabecera.AbsolutePosition & " de " & frmPropiedad.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmPropiedadDetalle = Nothing
    Cn.Execute "drop table " & sTemporal
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

Private Sub grdGrilla_DblClick()
   Dim nPos As Variant
   nPos = RsProducto.Bookmark
   Cn.Execute "update " & sTemporal & " set lEstado= " & IIf(RsProducto!lEstado, 0, 1) & " where tCodigoProducto='" & RsProducto!tCodigoProducto & "'"
   RsProducto.Requery
   RsProducto.Bookmark = nPos
End Sub

Public Sub ActivaPrecio(tActiva As String)
    If tActiva = "" Then
       fraValor.Enabled = False
       txtPrecio.Enabled = False
    Else
       fraValor.Enabled = True
       txtPrecio.Enabled = True
    End If
End Sub

Public Sub ActivaInventario(Activa As Boolean)
    If Not lAlmacen Then
       Activa = False
    End If
    fraInventario.Enabled = Activa
    cmdBusca.Enabled = Activa
End Sub




Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtPrecio
End Sub

Private Sub txtPrecio_LostFocus()
   txtPrecio.Text = Format(txtPrecio.Text, "###,###,##0.00")
End Sub

Public Sub Iniciar()
    ActivarBotones (False)
    Blanquear Me
    chkActivo.Value = 1
    chkCantidad.Value = 0
    txtPrecio.Text = "0.00"
    Cn.Execute "delete from " & sTemporal
    RsProducto.Requery
    txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    txtCodigoStock.Caption = ""
    txtDescripcionStock.Caption = ""
    txtDetallado.Text = ""
    cboArea.BoundText = "000"
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub
