VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmBusquedaSocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmBusquedaSocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtCodigoTarjeta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6840
      Width           =   6495
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   3360
      Left            =   30
      TabIndex        =   18
      Top             =   1410
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5927
      _LayoutType     =   4
      _RowHeight      =   28
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Listado"
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
      _StyleDefs(54)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
   Begin VB.TextBox txtNombres 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   950
      Width           =   6855
   End
   Begin VB.TextBox txtApellidos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   500
      Width           =   6855
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   50
      Width           =   6855
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   10
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4230
      Width           =   525
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   6
      ToolTipText     =   "Buscar Por Apellidos"
      Top             =   950
      Width           =   2655
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   3
      ToolTipText     =   "Buscar Por Código"
      Top             =   50
      Width           =   2655
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   5
      ToolTipText     =   "Buscar Por Apellidos"
      Top             =   500
      Width           =   2655
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Buscar Por Apellidos"
      Top             =   500
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Buscar Por Nombres"
      Top             =   950
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Buscar Por Código"
      Top             =   50
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpiarInvitado 
      Caption         =   "&Quitar"
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
      Left            =   9840
      TabIndex        =   8
      ToolTipText     =   "Quitar Invitado"
      Top             =   6285
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   8040
      TabIndex        =   20
      Top             =   4680
      Width           =   3570
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   9
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3675
      Width           =   525
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   8
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   525
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   7
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":1EA0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2565
      Width           =   525
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   6
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":276A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2010
      Width           =   525
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   5
      Left            =   8235
      Picture         =   "frmBusquedaSocio.frx":3034
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1455
      Width           =   525
   End
   Begin VB.CommandButton cmdInvitado 
      Caption         =   "&Asignar"
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
      Left            =   8040
      TabIndex        =   7
      ToolTipText     =   "Asignar Invitado"
      Top             =   6285
      Width           =   1725
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Estado del Titular :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   30
      Top             =   5550
      Width           =   3495
   End
   Begin VB.Label txtEstado 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   8160
      TabIndex        =   29
      Top             =   5850
      Width           =   3405
   End
   Begin VB.Image imgSFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   8805
      Picture         =   "frmBusquedaSocio.frx":38FE
      Stretch         =   -1  'True
      ToolTipText     =   "Foto"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   8805
      Stretch         =   -1  'True
      ToolTipText     =   "Foto"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label txtCategoria 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   8160
      TabIndex        =   28
      Top             =   5150
      Width           =   3405
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Categoria del Titular :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   27
      Top             =   4830
      Width           =   3495
   End
   Begin VB.Label txtSocioTitular 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   1560
      TabIndex        =   26
      Top             =   4875
      Width           =   6465
   End
   Begin VB.Label txtPariente 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   1560
      TabIndex        =   25
      Top             =   5595
      Width           =   6465
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Invitado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Pariente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Socio Titular :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label txtInvitado 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1560
      TabIndex        =   19
      Top             =   6315
      Width           =   6465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   21
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "frmBusquedaSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wAdiciona As Boolean
Public RsGrilla As Recordset

Dim sCriterio As String
Dim RsCliente As Recordset
Dim RsDireccion As Recordset
Dim RsDatos As Recordset

Public nPredeterm As Integer
Dim lConectado As Boolean

'invitado2013
Dim rsDatosInvitado As Recordset
Dim lPariente As Boolean


Public sCodigoPariente As String
Public sCodigoSocio As String
Public sCodigoInvitadoB As String

Dim fCodigo As String
Dim fApellido As String
Dim fNombre As String

Private Sub cmdBuscar_Click(Index As Integer)
Select Case Index
    Case 0 ' codigo
    
                sTipo = "Numero"
                frmNumPad.txtResultado.Caption = txtCodigo.Text
                frmNumPad.Show vbModal
                txtCodigo.Text = IIf(wEnter = True, sDescrip, txtCodigo.Text)
             
    Case 1 ' apellidos
    
                frmKeyBoard.txtResultado.Text = txtApellidos.Text
                frmKeyBoard.Show vbModal
                If wEnter Then
                   txtApellidos.Text = sDescrip
                End If
                wEnter = False
                
    Case 2 ' nombres
                
                frmKeyBoard.txtResultado.Text = txtNombres.Text
                
                frmKeyBoard.Show vbModal
                
                If wEnter Then
                   txtNombres.Text = sDescrip
                End If
                
                wEnter = False
                
    Case 3 ' quitar filtro
                    
                txtCodigo.Text = ""
                txtApellidos.Text = ""
                txtNombres.Text = ""
                txtSocioTitular.Caption = ""
                txtPariente.Caption = ""
                txtInvitado.Caption = ""
                txtEstado.Caption = ""
                txtCategoria.Caption = ""
                sCodigoPariente = ""
                sCodigoClienteFrecuente = ""
                sCodigoInvitado = ""
                imgFoto.Visible = False
                imgSFoto.Visible = True
                
                Isql = "usp_Inforest_ObtenerClientesFrecuente '9999','9999','9999'"
            
                Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
                
                Set grdGrilla.DataSource = RsGrilla
        
    Case 4 ' filtrar
       
            If txtCodigo.Text = "" And txtApellidos.Text = "" And txtNombres.Text = "" Then
                    MsgBox "Debe ingresar al menos un Criterio para Filtrar", vbInformation
                Exit Sub
            End If
            
            fCodigo = Trim(txtCodigo.Text)
            fApellido = Trim(txtApellidos.Text)
            fNombre = Trim(txtNombres.Text)
            
            Isql = "usp_Inforest_ObtenerClientesFrecuente '" & fCodigo & "','" & fApellido & "','" & fNombre & "'"
            
            Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
            
            Set grdGrilla.DataSource = RsGrilla
            
    Case 5 ' seleccionar
             If sModulo = "ADICION" Then
                    frmMozoUsuario.ReseteaTimer
             End If
            If txtSocioTitular.Caption = "" Then
                Exit Sub
            End If
            If RsGrilla.RecordCount = 0 Then
               Exit Sub
            End If
      
                If RsGrilla.RecordCount = 0 Then
                   Exit Sub
                End If
                 If Mid(RsGrilla.Fields("codigo"), 1, 1) = "C" Then
                        sCodigoParienteSeleccionado = Mid(RsGrilla!codigo, 2, 7)
                        Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla.Fields("coddely") & "'", Cn)
                Else
                        sCodigoParienteSeleccionado = ""
                        Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla!coddely & "'", Cn)
                End If
                
                'VALIDACION ACTIVO EL SOCIO TITULAR
                Dim lActivoDelivery As Boolean
                lActivoDelivery = Calcular("select isnull(lActivo,0) as codigo from vDelivery where codigo='" & RsGrilla!coddely & "'", Cn)
                   
                If lActivoDelivery = False Then
                   MsgBox "Por favor comuníquese al Dpto.de atención al Socio para confirmar sus Datos.", vbInformation
                   Exit Sub
                End If
                '------------------
                
                   
                '0612VALICACION
                If Not validaEstadoClienteFrecuente(RsDatos!tEstadoFrecuente) Then
                    Exit Sub
                End If
                '0612VALIDACION
                
                wEnter = True
                scodigo = IIf(RsGrilla.EOF = True, "", RsDatos!codigo)
                sDescrip = txtSocioTitular.Caption
                  nVar1 = IIf(IsNull(RsDatos!nDescuento), 0, Format(RsDatos!nDescuento, "##0.00"))
                'sVar1 = sTienda
                Unload Me
   
End Select
End Sub

'invitado2013
Private Sub cmdInvitado_Click()
    If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
    End If
    If txtSocioTitular.Caption = "" Then
       Exit Sub
    End If
    If RsGrilla.RecordCount = 0 Then
       Exit Sub
    End If
    Isql = "select * from vInvitado where codigodelivery like '" & sCodigoClienteFrecuente & "' Order by Invitado"
    Call ConfGrilla(2, frmBusquedaRapidaInvitado.grdGrilla, "Código", 2, "Codigo", 2000, 2, 0, "", _
                                                    "Invitado", 2, "Invitado", 6000, 0, 0, "")

    sTemp = txtInvitado.Caption
    frmBusquedaRapidaInvitado.nPredeterm = 1
    frmBusquedaRapidaInvitado.scodigoDelivery = RsGrilla!codigo
    frmBusquedaRapidaInvitado.Show vbModal
    If wEnter = True Then
        sCodigoInvitado = scodigo
        txtInvitado.Caption = sDescrip
    End If
End Sub
'invitado2013


Private Sub cmdLimpiarInvitado_Click()
    txtInvitado.Caption = ""
    sCodigoInvitado = ""
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   On Error GoTo err
   If (RsGrilla.EOF Or RsGrilla.BOF) Then
        Exit Sub
   End If
   
   Select Case Index
         
          
          Case Is = 5  ' Primero
               MoverPuntero Primero, grdGrilla
              ' txtTempo.SetFocus

          Case Is = 6  ' PG Up
               MoverPuntero pgup, grdGrilla
               'txtTempo.SetFocus

          Case Is = 7  ' Previo
               MoverPuntero previo, grdGrilla
               'txtTempo.SetFocus

          Case Is = 8  ' Siguiente
               MoverPuntero siguiente, grdGrilla
               'txtTempo.SetFocus

          Case Is = 9  ' PG Dn
               MoverPuntero pgdn, grdGrilla
               'txtTempo.SetFocus

          Case Is = 10 ' Ultimo
               MoverPuntero Ultimo, grdGrilla
               'txtTempo.SetFocus
   End Select
   Exit Sub
err:
   

End Sub

Private Sub Form_Activate()
 '  txtPredeterm.Caption = grdGrilla.Columns(nPredeterm).Caption
End Sub

Private Sub Form_Load()

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Isql = ""
    lPariente = False
   Centrar Me
   wEnter = False
   scodigo = ""
   sDescrip = ""
    
   Call ConfGrilla(5, grdGrilla, "Codigo", 2, "ttelefono", 1100, 2, 0, "", _
                                 "Cliente", 2, "Cliente", 5200, 0, 0, "", _
                                 "Identidad", 2, "tidentidad", 1200, 0, 0, "", _
                                 "CodigoSocio", 2, "codDely", 0, 0, 0, "", _
                                    "CodigoInt", 2, "codigo", 0, 0, 0, "")
      
   
   
   'Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
   
   'txtResultado.Caption = sTemp
   
   Screen.MousePointer = vbDefault
   
   
  End Sub

Private Sub Form_Unload(Cancel As Integer)
  
   Set RsCliente = Nothing
   Set RsGrilla = Nothing
   Set RsDireccion = Nothing
   Set frmBusquedaSocio = Nothing
End Sub

Private Sub grdGrilla_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
End Sub

Private Sub grdGrilla_DblClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
    If txtSocioTitular.Caption = "" Then
        Exit Sub
    End If

   If RsGrilla.RecordCount = 0 Then
      Exit Sub
   End If
     'If RsGrilla.RecordCount = 0 Or txtResultado.Caption = "" Then
                If RsGrilla.RecordCount = 0 Then
                   Exit Sub
                End If
                 If Mid(RsGrilla.Fields("codigo"), 1, 1) = "C" Then
                        sCodigoParienteSeleccionado = Mid(RsGrilla!codigo, 2, 7)
                        Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla.Fields("coddely") & "'", Cn)
                Else
                        sCodigoParienteSeleccionado = ""
                        Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla!coddely & "'", Cn)
                End If
                
                'VALIDACION ACTIVO EL SOCIO TITULAR
                Dim lActivoDelivery As Boolean
                lActivoDelivery = Calcular("select isnull(lActivo,0) as codigo from vDelivery where codigo='" & RsGrilla!coddely & "'", Cn)
                   
                If lActivoDelivery = False Then
                   MsgBox "Por favor comuníquese al Dpto.de atención al Socio para confirmar sus Datos.", vbInformation
                   Exit Sub
                End If
                '------------------
                
                
                '0612VALICACION
                If Not validaEstadoClienteFrecuente(RsDatos!tEstadoFrecuente) Then
                    Exit Sub
                End If
                '0612VALIDACION
                
                wEnter = True
                scodigo = IIf(RsGrilla.EOF = True, "", RsDatos!codigo)
                sDescrip = txtSocioTitular.Caption
                  nVar1 = IIf(IsNull(RsDatos!nDescuento), 0, Format(RsDatos!nDescuento, "##0.00"))
                'sVar1 = sTienda
                Unload Me
   
   
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo err
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If

    Screen.MousePointer = vbHourglass
    
    If Right(RsGrilla.Sort, 3) = "ASC" Then
       RsGrilla.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    Else
       RsGrilla.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    End If
    RsGrilla.Requery
    
    Screen.MousePointer = vbDefault
    nPredeterm = ColIndex
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    nPredeterm = ColIndex
    
End Sub

Public Sub Asigna()
   lPariente = False
   If RsGrilla.RecordCount > 0 Then
      If Mid(RsGrilla.Fields("codigo"), 1, 1) = "C" Then
            Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & LTrim(RsGrilla.Fields("codDely")) & "'", Cn)
            txtPariente.Caption = Trim(RsGrilla.Fields("cliente")) + " (" + RsGrilla.Fields("zona") + ")"
            sCodigoPariente = Mid(RsGrilla.Fields("codigo"), 2, Len(RsGrilla.Fields("codigo")))
            sCodigoClienteFrecuente = RsGrilla!coddely
             txtSocioTitular.Caption = Trim(IIf(IsNull(RsGrilla!tReferencia), "", RsGrilla!tReferencia))
            lPariente = True
        Else
            lPariente = False
            Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla!codigo & "'", Cn)
            txtPariente.Caption = ""
            sCodigoPariente = ""
            sCodigoClienteFrecuente = RsGrilla!codigo
            txtSocioTitular.Caption = Trim(IIf(IsNull(RsDatos!tApellido), "", RsDatos!tApellido)) + " " + Trim(IIf(IsNull(RsDatos!tNombre), "", RsDatos!tNombre))
        
      End If
      sCodigoInvitado = ""
      txtInvitado.Caption = ""
      txtEstado.Caption = IIf(IsNull(RsDatos.Fields("estadofrecuente")), "", RsDatos.Fields("estadofrecuente"))
      txtCategoria.Caption = IIf(IsNull(RsDatos.Fields("tipocliente")), "", RsDatos.Fields("tipocliente"))
      'invitado2013
      Set rsDatosInvitado = Lib.OpenRecordset("select * from tdeliveryinvitado where tcodigoinvitado='" & sCodigoInvitado & "'", Cn)
      
      'invitado2013
          'invitado2013
      If rsDatosInvitado.RecordCount > 0 Then
            txtInvitado = IIf(IsNull(rsDatosInvitado!tNombre), "", rsDatosInvitado!tNombre) & " " & IIf(IsNull(rsDatosInvitado!tApellido), "", rsDatosInvitado!tApellido)
      Else
            txtInvitado.Caption = ""
      End If
        'invitado2013
          'invitado2013
        'fotoooooooooo
       On Error GoTo err
       Dim rst1 As New ADODB.Recordset
       imgFoto.DataField = "foto"
       If lPariente = False Then
                    
                    
                    Set rst1 = Lib.OpenRecordset("select ifoto as foto from tdelivery where tcodigodelivery='" & sCodigoClienteFrecuente & "'", Cn)
                    If Not (rst1.EOF Or rst1.BOF) Then
                        If IsNull(rst1.Fields(0)) Then
                            imgFoto.Visible = False
                            imgSFoto.Visible = True
                        Else
                            Set imgFoto.DataSource = rst1
                            imgFoto.Visible = True
                            imgSFoto.Visible = False
                        End If
                    Else
                        imgFoto.Visible = False
                        imgSFoto.Visible = True
                    End If
        Else
        
                    Set rst1 = Lib.OpenRecordset("select ifoto as foto from tpariente where tcodigopariente='" & sCodigoPariente & "'", Cn)
                    If Not (rst1.EOF Or rst1.BOF) Then
                        If IsNull(rst1.Fields(0)) Then
                            imgFoto.Visible = False
                            imgSFoto.Visible = True
                        Else
                            Set imgFoto.DataSource = rst1
                            imgFoto.Visible = True
                            imgSFoto.Visible = False
                        End If
                    Else
                        imgFoto.Visible = False
                        imgSFoto.Visible = True
                    End If
        
        End If
       
   Else
   '  txtResultado.Caption = ""
     txtSocioTitular.Caption = ""
     txtPariente.Caption = ""
     'invitado2013
     txtInvitado.Caption = ""
     txtEstado.Caption = ""
     txtCategoria.Caption = ""
     'invitado2013
   End If
   Set RsDatos = Nothing
'   txtTempo.SetFocus
   Exit Sub
err:
    imgFoto.Picture = Nothing
    imgFoto.Visible = False
    imgSFoto.Visible = True
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
    
  ' txtTempo.SetFocus
   Limpiar
   If Not RsGrilla.EOF Then
      Asigna
   End If
   Exit Sub
err:
   
End Sub
 
Public Sub Limpiar()
    txtSocioTitular.Caption = ""
    txtEstado.Caption = ""
    txtCategoria.Caption = ""
    txtPariente.Caption = ""
    'txtInvitado.Caption = ""
    sCodigoPariente = ""
    sCodigoClienteFrecuente = ""
  ' sCodigoInvitado = ""
   '  txtTempo.SetFocus
End Sub
 
Public Function validaEstadoClienteFrecuente(tcodigoEstado As String) As Boolean
    validaEstadoClienteFrecuente = True
    Dim lAlerta As Boolean
    Dim lRestriccion As Boolean

    lAlerta = Calcular("select isnull(nboton,0) as codigo from vestadofrecuente where codigo='" & tcodigoEstado & "'", Cn)
    lRestriccion = Calcular("select isnull(nvalor,0) as codigo from vestadofrecuente where codigo='" & tcodigoEstado & "'", Cn)
    
    
    If lAlerta = True Then
        MsgBox "   -   -   ALERTA   -   -   " & vbCrLf & " EL ESTADO DEL CLIENTE SELECCIONADO ES: " & UCase(Calcular("select tresumido as codigo from vestadofrecuente where codigo='" & tcodigoEstado & "'", Cn)), vbInformation
    End If
    
    If lRestriccion = True Then
        MsgBox "  -  -  ALERTA : RESTRICCION -  -  " & vbCrLf & " EL ESTADO DEL CLIENTE SELECCIONADO ES: " & UCase(Calcular("select tresumido as codigo from vestadofrecuente where codigo='" & tcodigoEstado & "'", Cn)), vbCritical
        validaEstadoClienteFrecuente = False
    End If
    
End Function

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    cmdBuscar_Click (4)
  End If
End Sub

Private Sub txtApellidos_LostFocus()
 Call ValidaStr(txtApellidos)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
 '  TabNext KeyAscii
  If KeyAscii = 13 Then
    cmdBuscar_Click (4)
  End If
   Numerico KeyAscii, txtCodigo
     
End Sub

Private Sub txtCodigoTarjeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigoTarjeta.Text <> "" Then
        Dim xCodigoTarjeta As String
        Dim xCodigo As String
        xCodigoTarjeta = txtCodigoTarjeta.Text
        xCodigo = Mid(xCodigoTarjeta, 110, 6) '+ "-" + Mid(xCodigoTarjeta, 116, 2)
        txtCodigo.Text = txtCodigo.Text + xCodigo
        cmdBuscar_Click (4)
        txtCodigo.Text = ""
        txtCodigoTarjeta.Text = ""
        txtCodigoTarjeta.SetFocus
     End If
End Sub


Private Sub txtNombres_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdBuscar_Click (4)
  End If
End Sub

Private Sub txtNombres_LostFocus()
 Call ValidaStr(txtNombres)
End Sub
