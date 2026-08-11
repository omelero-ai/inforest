VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmVentaPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11910
   ForeColor       =   &H00808080&
   Icon            =   "frmVentaPV.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPedido 
      Caption         =   " Pedido "
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
      Height          =   3555
      Left            =   0
      TabIndex        =   212
      Top             =   0
      Width           =   2625
      Begin VB.Label txtReferencia 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   930
         TabIndex        =   226
         Top             =   2190
         Width           =   1620
      End
      Begin VB.Label txtDireccion 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   630
         Left            =   930
         TabIndex        =   225
         Top             =   2850
         Width           =   1620
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   630
         Left            =   930
         TabIndex        =   224
         Top             =   1500
         Width           =   1620
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   930
         TabIndex        =   223
         Top             =   2520
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   255
         TabIndex        =   222
         Top             =   2573
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   225
         TabIndex        =   221
         Top             =   3083
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Guia :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   540
         TabIndex        =   220
         Top             =   1193
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   390
         TabIndex        =   219
         Top             =   1733
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   435
         TabIndex        =   218
         Top             =   413
         Width           =   450
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   930
         TabIndex        =   217
         Top             =   1140
         Width           =   1620
      End
      Begin VB.Label txtFecha 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   930
         TabIndex        =   216
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label txtTipoPedido 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   930
         TabIndex        =   215
         Top             =   795
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   60
         TabIndex        =   214
         Top             =   848
         Width           =   825
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tienda :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   405
         TabIndex        =   213
         Top             =   2250
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Visualizar Pedido"
      Height          =   585
      Index           =   13
      Left            =   6510
      TabIndex        =   149
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Código Directo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   7
      Left            =   6510
      TabIndex        =   148
      Top             =   7140
      Width           =   1185
   End
   Begin VB.Frame fraEmisiones 
      Caption         =   " Emisiones Rápidas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   900
      Left            =   45
      TabIndex        =   103
      Top             =   7740
      Width           =   7665
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cuenta Corritentes"
         Height          =   555
         Index           =   14
         Left            =   5092
         TabIndex        =   160
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Pre Cuenta"
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
         Left            =   6345
         Picture         =   "frmVentaPV.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   3
         Left            =   2520
         TabIndex        =   106
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   2
         Left            =   1305
         TabIndex        =   105
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   1
         Left            =   90
         TabIndex        =   104
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   4
         Left            =   3735
         TabIndex        =   202
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Precios"
      Height          =   585
      Index           =   6
      Left            =   6510
      TabIndex        =   26
      Top             =   6555
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Propiedad y Observación"
      Height          =   585
      Index           =   4
      Left            =   6510
      TabIndex        =   25
      Top             =   5970
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Disminuir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   6510
      Picture         =   "frmVentaPV.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5385
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Aumentar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmVentaPV.frx":09BE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Cantidad"
      Height          =   585
      Index           =   1
      Left            =   6510
      TabIndex        =   22
      Top             =   4215
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Elimina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmVentaPV.frx":0AC0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3630
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Transferencia"
      Height          =   585
      Index           =   4
      Left            =   6510
      TabIndex        =   20
      Top             =   1815
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Elimina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmVentaPV.frx":0BC2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1230
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Modifica"
      Height          =   585
      Index           =   1
      Left            =   6510
      TabIndex        =   18
      Top             =   645
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmVentaPV.frx":0CC4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   11
      Left            =   5880
      Picture         =   "frmVentaPV.frx":11F6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   10
      Left            =   5880
      Picture         =   "frmVentaPV.frx":1AC0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5970
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   9
      Left            =   5880
      Picture         =   "frmVentaPV.frx":238A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5385
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   8
      Left            =   5880
      Picture         =   "frmVentaPV.frx":2C54
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   7
      Left            =   5880
      Picture         =   "frmVentaPV.frx":351E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   6
      Left            =   5880
      Picture         =   "frmVentaPV.frx":3DE8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3630
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   5
      Left            =   5880
      Picture         =   "frmVentaPV.frx":46B2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2985
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   4
      Left            =   5880
      Picture         =   "frmVentaPV.frx":4F7C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   3
      Left            =   5880
      Picture         =   "frmVentaPV.frx":5846
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1815
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   2
      Left            =   5880
      Picture         =   "frmVentaPV.frx":6110
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1230
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   1
      Left            =   5880
      Picture         =   "frmVentaPV.frx":69DA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   645
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   0
      Left            =   5880
      Picture         =   "frmVentaPV.frx":72A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
   Begin TrueOleDBGrid80.TDBGrid grdDetalle 
      Height          =   3750
      Left            =   30
      TabIndex        =   159
      Top             =   3600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6615
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
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   0
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
   Begin VB.TextBox txtBarra 
      Height          =   345
      Left            =   5250
      TabIndex        =   1
      Top             =   8040
      Width           =   1785
   End
   Begin VB.Frame fraDetalle 
      Caption         =   " Detalle del Producto "
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
      Height          =   6585
      Left            =   7740
      TabIndex        =   107
      Top             =   90
      Width           =   4155
      Begin VB.CommandButton cmdCortesia 
         Caption         =   "Cortesía"
         Height          =   585
         Left            =   180
         TabIndex        =   147
         Top             =   990
         Width           =   1155
      End
      Begin VB.CommandButton cmdPrecio 
         Caption         =   "Precio"
         Height          =   585
         Left            =   180
         TabIndex        =   146
         Top             =   300
         Width           =   1155
      End
      Begin VB.Frame fraImpuesto 
         Caption         =   " Impuestos "
         ForeColor       =   &H00800080&
         Height          =   975
         Left            =   120
         TabIndex        =   142
         Top             =   1695
         Width           =   3960
         Begin VB.CommandButton cmdImpuesto 
            Height          =   585
            Index           =   2
            Left            =   2730
            TabIndex        =   145
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdImpuesto 
            Height          =   585
            Index           =   1
            Left            =   1440
            TabIndex        =   144
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdImpuesto 
            Height          =   585
            Index           =   0
            Left            =   150
            TabIndex        =   143
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.Frame fraPrecio 
         Caption         =   " Precio de Venta "
         ForeColor       =   &H00800080&
         Height          =   3795
         Left            =   90
         TabIndex        =   113
         Top             =   2700
         Width           =   3990
         Begin VB.Label txtCantidad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   141
            Top             =   2580
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cantidad :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   375
            TabIndex        =   140
            Top             =   2625
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Recargo :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   390
            TabIndex        =   139
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   120
            TabIndex        =   138
            Top             =   2340
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descuento :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   240
            TabIndex        =   137
            Top             =   585
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Oficial :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   90
            TabIndex        =   136
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   240
            TabIndex        =   135
            Top             =   1725
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   240
            TabIndex        =   134
            Top             =   1455
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 3 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   240
            TabIndex        =   133
            Top             =   2010
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   18
            Left            =   600
            TabIndex        =   132
            Top             =   2925
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Observación :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   131
            Top             =   3165
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   9
            Left            =   165
            TabIndex        =   130
            Top             =   1185
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   19
            Left            =   1590
            TabIndex        =   129
            Top             =   570
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   20
            Left            =   1590
            TabIndex        =   128
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cortesía"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   2490
            TabIndex        =   127
            Top             =   2325
            Width           =   555
         End
         Begin VB.Label txtImpuesto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   126
            Top             =   1410
            Width           =   1365
         End
         Begin VB.Label txtImpuesto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   125
            Top             =   1695
            Width           =   1365
         End
         Begin VB.Label txtImpuesto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   124
            Top             =   1965
            Width           =   1365
         End
         Begin VB.Label txtDPorcentaje 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1065
            TabIndex        =   123
            Top             =   540
            Width           =   510
         End
         Begin VB.Label txtRPorcentaje 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1065
            TabIndex        =   122
            Top             =   810
            Width           =   510
         End
         Begin VB.Label txtRImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1785
            TabIndex        =   121
            Top             =   810
            Width           =   645
         End
         Begin VB.Label txtDImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1785
            TabIndex        =   120
            Top             =   540
            Width           =   645
         End
         Begin VB.Label txtNeto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   1065
            TabIndex        =   119
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label txtPVenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   118
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label txtOficial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   117
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label txtVenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   116
            Top             =   2850
            Width           =   1365
         End
         Begin VB.Label txtCortesia 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   2475
            TabIndex        =   115
            Top             =   2550
            Width           =   1440
         End
         Begin VB.Label txtObserva 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   1065
            TabIndex        =   114
            Top             =   3165
            Width           =   2850
         End
      End
      Begin VB.Frame fraDescuento 
         Caption         =   " Descuentos / Recargos "
         ForeColor       =   &H00800080&
         Height          =   1515
         Left            =   1440
         TabIndex        =   108
         Top             =   180
         Width           =   2640
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "( % ) del Recargo"
            Height          =   555
            Index           =   3
            Left            =   1380
            TabIndex        =   112
            Top             =   870
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "( % ) del Descuento"
            Height          =   555
            Index           =   1
            Left            =   1380
            TabIndex        =   111
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "Monto del Recargo"
            Height          =   555
            Index           =   2
            Left            =   150
            TabIndex        =   110
            Top             =   870
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "Monto del Descuento"
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   109
            Top             =   270
            Width           =   1155
         End
      End
   End
   Begin VB.Frame fraPlato 
      Caption         =   " Platos de Venta "
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
      Height          =   6585
      Left            =   7740
      TabIndex        =   29
      Top             =   45
      Width           =   4155
      Begin VB.Frame fraSubGrupo 
         Caption         =   " SubGrupos "
         ForeColor       =   &H000000C0&
         Height          =   3375
         Left            =   60
         TabIndex        =   32
         Top             =   3000
         Width           =   1635
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   2460
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   840
            Picture         =   "frmVentaPV.frx":7B6E
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2460
            Width           =   720
         End
      End
      Begin VB.Frame fraGrupo 
         Caption         =   " Grupos "
         ForeColor       =   &H000000C0&
         Height          =   2595
         Left            =   60
         TabIndex        =   30
         Top             =   270
         Width           =   1635
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1770
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   0
            Left            =   840
            Picture         =   "frmVentaPV.frx":7FB0
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1770
            Width           =   720
         End
      End
      Begin VB.Frame fraProducto 
         Caption         =   " Productos "
         ForeColor       =   &H000000C0&
         Height          =   6105
         Left            =   1740
         TabIndex        =   31
         Top             =   270
         Width           =   2325
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   90
            TabIndex        =   67
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   810
            TabIndex        =   66
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1530
            TabIndex        =   65
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   90
            TabIndex        =   64
            Top             =   1140
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   810
            TabIndex        =   63
            Top             =   1140
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   1530
            TabIndex        =   62
            Top             =   1140
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   90
            TabIndex        =   61
            Top             =   1950
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   810
            TabIndex        =   60
            Top             =   1950
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   1530
            TabIndex        =   59
            Top             =   1950
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   90
            TabIndex        =   58
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   810
            TabIndex        =   57
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   1530
            TabIndex        =   56
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   90
            TabIndex        =   55
            Top             =   3570
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   810
            TabIndex        =   54
            Top             =   3570
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   1530
            TabIndex        =   53
            Top             =   3570
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   90
            TabIndex        =   52
            Top             =   4380
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   810
            TabIndex        =   51
            Top             =   4380
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1530
            TabIndex        =   50
            Top             =   4380
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   90
            TabIndex        =   49
            Top             =   5190
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   20
            Left            =   810
            TabIndex        =   48
            Top             =   5190
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   1530
            Picture         =   "frmVentaPV.frx":83F2
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   5190
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraEliminacion 
      Caption         =   " Motivo de Eliminación "
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
      Height          =   6585
      Left            =   7740
      TabIndex        =   162
      Top             =   45
      Width           =   4155
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1710
         TabIndex        =   201
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   930
         TabIndex        =   200
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   150
         TabIndex        =   199
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3300
         TabIndex        =   198
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   150
         TabIndex        =   197
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   930
         TabIndex        =   196
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1710
         TabIndex        =   195
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2505
         TabIndex        =   194
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3300
         TabIndex        =   193
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   150
         TabIndex        =   192
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   930
         TabIndex        =   191
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1725
         TabIndex        =   190
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2505
         TabIndex        =   189
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3300
         TabIndex        =   188
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   150
         TabIndex        =   187
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   930
         TabIndex        =   186
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1725
         TabIndex        =   185
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2505
         TabIndex        =   184
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3300
         TabIndex        =   183
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   150
         TabIndex        =   182
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   930
         TabIndex        =   181
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1725
         TabIndex        =   180
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2505
         TabIndex        =   179
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3300
         TabIndex        =   178
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   150
         TabIndex        =   177
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   930
         TabIndex        =   176
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1725
         TabIndex        =   175
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2505
         TabIndex        =   174
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3300
         TabIndex        =   173
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   150
         TabIndex        =   172
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   930
         TabIndex        =   171
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1725
         TabIndex        =   170
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2505
         TabIndex        =   169
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3300
         TabIndex        =   168
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   150
         TabIndex        =   167
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   930
         TabIndex        =   166
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1725
         TabIndex        =   165
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2505
         TabIndex        =   164
         Top             =   1080
         Width           =   720
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
         Height          =   720
         Index           =   17
         Left            =   2505
         Picture         =   "frmVentaPV.frx":8834
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   5625
         Width           =   1515
      End
   End
   Begin VB.Frame fraPropiedad 
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
      Height          =   8610
      Left            =   7740
      TabIndex        =   68
      Top             =   -270
      Width           =   4155
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   3345
         Picture         =   "frmVentaPV.frx":8DBE
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   4950
         Width           =   720
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   5415
         Left            =   1672
         TabIndex        =   204
         Top             =   225
         Width           =   60
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   870
         Picture         =   "frmVentaPV.frx":9200
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   1440
         Picture         =   "frmVentaPV.frx":9642
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   7785
         Width           =   1470
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
         Height          =   720
         Index           =   9
         Left            =   2970
         Picture         =   "frmVentaPV.frx":9784
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   7785
         Width           =   1110
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   2580
         TabIndex        =   101
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   1815
         TabIndex        =   100
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   3345
         TabIndex        =   97
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   2580
         TabIndex        =   96
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   1815
         TabIndex        =   95
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   3345
         TabIndex        =   92
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   2580
         TabIndex        =   91
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   1815
         TabIndex        =   90
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   3345
         TabIndex        =   87
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   2580
         TabIndex        =   86
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   1815
         TabIndex        =   85
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   3345
         TabIndex        =   82
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2580
         TabIndex        =   81
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1815
         TabIndex        =   80
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3345
         TabIndex        =   77
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   2580
         TabIndex        =   76
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   1815
         TabIndex        =   75
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   3345
         TabIndex        =   72
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   2580
         TabIndex        =   71
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   1815
         TabIndex        =   70
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4950
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Observacion :"
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
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   211
         Top             =   6930
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propiedad :"
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
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   210
         Top             =   5760
         Width           =   990
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Propiedad "
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
         Height          =   195
         Index           =   3
         Left            =   1890
         TabIndex        =   209
         Top             =   45
         Width           =   1050
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Operador   "
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
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   208
         Top             =   45
         Width           =   1095
      End
      Begin VB.Label lblObservacion 
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
         Height          =   555
         Left            =   90
         TabIndex        =   207
         Top             =   7155
         Width           =   3975
      End
      Begin VB.Label lblResumen 
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
         Height          =   870
         Left            =   90
         TabIndex        =   206
         Top             =   5985
         Width           =   3975
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opcion "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   7740
      TabIndex        =   150
      Top             =   6615
      Width           =   4155
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Importar Pedidos"
         Height          =   585
         Index           =   8
         Left            =   135
         TabIndex        =   152
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Caja Rápida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   16
         Left            =   2790
         Picture         =   "frmVentaPV.frx":9876
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   330
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Pagos y Div.Docum"
         Height          =   585
         Index           =   6
         Left            =   1470
         TabIndex        =   157
         Top             =   330
         Width           =   1230
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
         Height          =   585
         Index           =   0
         Left            =   2790
         Picture         =   "frmVentaPV.frx":99C0
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Busca Pedido"
         Height          =   585
         Index           =   10
         Left            =   135
         TabIndex        =   153
         Top             =   330
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Refrescar"
         Height          =   585
         Index           =   11
         Left            =   1470
         TabIndex        =   151
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         Height          =   585
         Index           =   2
         Left            =   1470
         TabIndex        =   155
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Aceptar"
         Height          =   585
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   1020
         Width           =   1230
      End
   End
   Begin TrueOleDBGrid80.TDBGrid grdCabecera 
      Height          =   3555
      Left            =   2640
      TabIndex        =   227
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   6271
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
      Splits(0).ScrollBars=   0
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
      Caption         =   "Cabecera de Pedidos"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "U.Fact :"
      Height          =   195
      Index           =   23
      Left            =   7740
      TabIndex        =   28
      Top             =   8400
      Width           =   570
   End
   Begin VB.Label lblFactura 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   8340
      TabIndex        =   27
      Top             =   8370
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      Caption         =   "10/01/1970"
      Height          =   255
      Left            =   9480
      TabIndex        =   16
      Top             =   8370
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   24
      Left            =   3780
      TabIndex        =   3
      Top             =   7470
      Width           =   1005
   End
   Begin VB.Label txtSuma 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4845
      TabIndex        =   2
      Top             =   7425
      Width           =   990
   End
   Begin VB.Label txtDescuento 
      Caption         =   "Label2"
      Height          =   240
      Left            =   450
      TabIndex        =   228
      Top             =   8145
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmVentaPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As ADODB.Recordset
Public RsDetalle As ADODB.Recordset
Dim RsImpresion As Recordset
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsProducto As Recordset
Dim RsArea As Recordset
Dim RsPropiedad As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsTipoDocumento As Recordset
Dim RsT3 As Recordset
Dim RsPC As Recordset
Dim RsOferta As Recordset
Dim RsMotivoEliminacion As Recordset
Dim RsOperador As Recordset

'Variables Generales
Dim wDetalle As Boolean
Dim wAgrega As Boolean

'Variables Cabecera
Dim sTipoPedido As String
Dim sTipoAtencion As String
Dim sCortesia As String
Dim nCorrela As Long
Dim ntTiempo As Integer
Dim sdetalleconsumo As String
Public lIncluido As Boolean

Dim sCompania As String
Dim sContacto As String
Dim sComandaInfhotel As String
Dim wCabecera As Boolean

'Variables Detalles
Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim nCantidad As Double
Dim sItem As String
Dim xItem As String
Dim sProducto As String
Dim sGrupo As String
Dim xGrupo As String
Dim sSubGrupo As String
Dim xSubGrupo As String
Dim sTD As String
Dim xsuma As Double
Dim xLinea As Double
Dim xConsumo As Double
Dim xDescuento As Double
Dim lPrecuenta As Boolean
Dim lPropiedad As Boolean
Dim sMonedaBase As String
Dim sDescuento As Double
Dim Acumulado As Double
Dim sUsuarioAutoriza As String
Dim sCodigoDescuento As String
Dim sDescripcionDescuento As String
Dim tAutorizaDescuento As String
Dim wCalcula As Boolean
Dim UltimaComanda As String
Dim PrecioTotal As Double
Dim EstadoC As String
Private Sub cmdBuscar_Click(Index As Integer)

    Dim i As Integer
    Sw = False
    Select Case Index
           Case Is = 0 'Grupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                Isql = "select * from vGrupo where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Grupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                   For i = 1 To 5
                       cmdGrupo(i).BackColor = vbButtonFace
                       cmdSubGrupo(i).BackColor = vbButtonFace
                   Next i
                   cmdSubGrupo(6).BackColor = vbButtonFace
                   cmdSubGrupo(7).BackColor = vbButtonFace
                   
                   sGrupo = sCodigo
                   xGrupo = sCodigo
                   RsGrupo.MoveFirst
                   RsGrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
                      cmdGrupo(RsGrupo!nBoton).BackColor = vbBlue
                   End If
                   RsSubGrupo.Filter = ("tGrupo = '" & xGrupo & "'")
                   Call AsignaBoton(7, RsSubGrupo, cmdSubGrupo())
                                      
                   If RsSubGrupo.RecordCount <> 0 Then
                      RsSubGrupo.MoveFirst
                      RsSubGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "' and lActivo=1", Cn) & "'"
                      sSubGrupo = RsSubGrupo!Codigo
                      xSubGrupo = RsSubGrupo!Codigo
                      If Not RsSubGrupo.EOF And RsSubGrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubGrupo!nBoton).BackColor = vbRed
                      End If
                   Else
                      sSubGrupo = ""
                      xSubGrupo = ""
                   End If
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido)
                End If
                txtBarra.SetFocus
           
           Case Is = 1 'SubGrupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                Isql = "select * from vSubGrupo where lActivo = 1 and tGrupo ='" & xGrupo & "' Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "SubGrupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                         
                If wEnter Then
                   For i = 1 To 7
                       cmdSubGrupo(i).BackColor = vbButtonFace
                   Next i
                
                   sSubGrupo = sCodigo
                   xSubGrupo = sCodigo
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   
                   RsSubGrupo.MoveFirst
                   RsSubGrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsSubGrupo.EOF And RsSubGrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubGrupo!nBoton).BackColor = vbRed
                   End If
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido)
                End If
                txtBarra.SetFocus
           
           Case Is = 2 'Producto
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                Isql = "select * from vProducto where lActivo = 1 and tSubGrupo ='" & xSubGrupo & "' and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
                Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                                "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                                "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                                "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                                "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
   
                If wEnter Then
                   sProducto = sCodigo
                   RsProducto.MoveFirst
                   RsProducto.Find "Codigo = '" & sProducto & "'"
                   If vComanda Then
                      sTipo = "Comanda"
                      sDescrip = sComanda
                      frmNumPad.Caption = "Comanda"
                      frmNumPad.Show vbModal
                      sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                      
                     If wComanda And sComanda = "" Then
                        MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                        Exit Sub
                     End If
                   
                     If Not wEnter Then
                        Exit Sub
                     End If
                   End If
                   
                   InsertaProducto 1
                   
                   If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                      lPropiedad = True
                   End If
                End If
                txtBarra.SetFocus
                
           Case Is = 6 'Operadores
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTipo = ""
                sTemp = ""
                Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                    Screen.MousePointer = vbHourglass
                    For i = 1 To 13
                        cmdOperador(i).BackColor = vbButtonFace
                    Next i
                    RsOperador.MoveFirst
                    RsOperador.Find "Codigo='" & sCodigo & "'"
                    If Not RsOperador.EOF And RsOperador!nBoton > 0 Then
                       cmdOperador(RsOperador!nBoton).BackColor = vbRed
                    End If
                    AsignaPropiedad
                    Screen.MousePointer = vbDefault
                End If
                txtBarra.SetFocus

           Case Is = 7 'Propiedades
                Exit Sub   'Pensarlo mas adelante
'                sTipo = ""
'                sTemp = ""
'                Isql = "select * from vPropiedades where lActivo = 1 and tOperador='" & RsOperador!Codigo & "' Order by Descripcion "
'                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
'                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
'                frmBusquedaRapida.nPredeterm = 1
'                frmBusquedaRapida.Show vbModal
'
'                If wEnter = True Then
'                End If
'                txtBarra.SetFocus
   
   End Select
End Sub

Private Sub cmdCabecera_Click(Index As Integer)
    Dim sEstado As String
    txtBarra.SetFocus
    Dim j As Integer
    wCalcula = False
    
    Select Case Index
           Case Is = 0 'Nuevo
                Sw = True
                frmNuevo.Show vbModal
                sPedido = ""
                
                If wEnter Then
                   ActivaOpcion True
                   ActivaFrame fraPlato
                   ActivaCabecera True
                                     
                   'Obtiene el Numero de Orden
                   nCorrela = Calcular("select nCORRELATIVO as Codigo from TPARAMETRO", Cn) + 1
                   If IsNull(nCorrela) Then
                       nCorrela = 1
                   End If
                   Cn.Execute "Update TPARAMETRO Set nCorrelativo = " & nCorrela
 
                   'Obtiene el Numero de Pedido
                   sPedido = Calcular("select max(tCodigoPedido) as Codigo from mpedido where substring(tCodigoPedido,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                   If IsNull(sPedido) Or Mid(sPedido, 1, 2) <> Mid(Trim(Str(Year(Date))), 3, 2) Then
                       sPedido = Mid(Trim(Str(Year(Date))), 3, 2) & "00000001"
                   Else
                       sPedido = Mid(Trim(Str(Year(Date))), 3, 2) & Lib.Correlativo(Mid(sPedido, 3, 8), 8)
                   End If
                                         
                   Isql = "insert into MPEDIDO " & _
                          "(tCodigoPedido,nCorrelativo, tClienteCorp, tTienda, fFecha,tEstadoPedido,tTipoPedido,lPrioridad, " & _
                          " tTipoAtencion,tCaja, tSalon, tTurno, tObservacion,tUsuario,fRegistro)" & _
                          "Values(   '" & sPedido & "'," _
                                        & nCorrela & "," _
                                        & "'" & sCliente & "'," _
                                        & "'" & sTienda & "'," _
                                        & "getdate() ," _
                                        & "'01'," _
                                        & "'04'," _
                                        & "'0'," _
                                        & "'01'," _
                                        & "'" & sCaja & "'," _
                                        & "'" & sSalon & "'," _
                                        & "'" & sTurno & "'," _
                                        & "'" & sTemp & "'," _
                                        & "'" & sUsuario & "'," _
                                        & "getDate() )"
                   Cn.Execute Isql
                  
                   RsCabecera.Requery
                   RsCabecera.MoveFirst
                   RsCabecera.Find "Codigo = '" & sPedido & "'"
                   If RsCabecera.RecordCount <> 0 Then
                      AsignaPedido
                   End If
                End If
                
                grdCabecera.Enabled = True
                'Llena el Detalle
                RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
                
           Case Is = 1 'Modifica
                If RsCabecera.RecordCount <> 0 Then
                   Sw = False
                   frmNuevo.txtSerie = Mid(txtObservacion.Caption, 3, 5)
                   frmNuevo.txtCorrelativo = Mid(txtObservacion.Caption, 9, 9)
                   frmNuevo.txtCliente = txtCliente.Caption
                   frmNuevo.txtReferencia = txtReferencia.Caption
                   frmNuevo.Show vbModal
                   
                   If wEnter Then
                      Isql = "Update MPEDIDO Set " & _
                             "tClienteCorp ='" & sCliente & "', " & _
                             "tTienda ='" & sTienda & "', " & _
                             "tObservacion ='" & sTemp & "' " & _
                             " where tCodigoPedido = '" & sPedido & "'"
                      Cn.Execute Isql
                      RsCabecera.Requery
                      RsCabecera.Find "Codigo = '" & sPedido & "'"
                      AsignaPedido
                   End If
                End If
                
           Case Is = 2 'Elimina
                Dim i As Integer
                If sPedido <> "" Then
                   If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0')<>'0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
                      MsgBox "Imposible eliminar pedidos con items Facturados", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                   If MsgBox("Seguro de Eliminar el Pedido?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                      sUsuarioAutoriza = sUsuario
                      If lPasswordC Then
                         If Supervisor("01") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuarioAutoriza = sVar1
                      End If
                                                                                                                                                                                         
                      If lEliminaC Then
                         wCabecera = True
                         ActivaFrame fraEliminacion
                         ActivaCabecera False
                         ActivaOpcion False
                         cmdOpcion(1).Enabled = False
                         cmdOpcion(2).Enabled = False
                      Else
                         sCodigo = ""
                         sDescrip = ""
                         EliminaCabecera
                      End If
                   End If
                End If
                
           Case Is = 4 'Transferencia
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                frmDivision.txtFecha = txtFecha.Caption
                frmDivision.txtTipoPedido.Caption = txtTipoPedido.Caption
                frmDivision.txtDescuento.Caption = "0.00"
                frmDivision.txtObservacion.Caption = txtObservacion.Caption
                frmDivision.Label1(7).Caption = "Cliente :"
                frmDivision.cmdOpcion(2).Enabled = False
                frmDivision.cmdOpcion(3).Enabled = False
                frmDivision.cmdOpcion(4).Enabled = False
                frmDivision.cmdDetalle(1).Enabled = False
                frmDivision.txtMotorizado.Caption = grdCabecera.Columns(0).Text
                
                frmDivision.Show vbModal
                If wEnter Then
                   RsCabecera.Requery
                   RsDetalle.Requery
                   RsProductoPropiedad.Requery
                   If RsCabecera.RecordCount <> 0 Then
                      RsCabecera.Find "Codigo = '" & sPedido & "'"
                   End If
                End If
                                                           
    End Select
End Sub

Private Sub cmdCortesia_Click()
    sTemp = ""
    Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    
    If wEnter = True Then
       sCortesia = sCodigo
       txtCortesia.Caption = sDescrip
       
       txtDImporte.Caption = "0.00"
       txtRImporte.Caption = "0.00"
       txtImpuesto1.Caption = "0.00"
       txtImpuesto2.Caption = "0.00"
       txtImpuesto3.Caption = "0.00"
       txtPVenta.Caption = "0.00"
       txtVenta.Caption = "0.00"
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nPBase = 0
       nRecargo = 0
       nDescuento = 0
       nPVenta = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
       sCortesia = ""
       txtCortesia.Caption = ""
    End If
End Sub

Private Sub cmdDescuento_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Dscto. Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nDescuento = Val(sDescrip)
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 1 ' Dscto. Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtDPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nDescuento = nOficial * Val(sDescrip) / 100
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 2 ' Recargo Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nRecargo = Val(sDescrip)
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 3 ' Recargo Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtRPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nRecargo = nOficial * Val(sDescrip) / 100
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
    End Select
End Sub

Private Sub cmdDetalle_Click(Index As Integer)
   Dim nPos As Integer
   txtBarra.SetFocus
   If RsDetalle.RecordCount = 0 Then
      Exit Sub
   End If
   If RsDetalle!tFacturado = "F" Or RsDetalle!tFacturado = "P" Then
      MsgBox "Producto ya Facturado ", vbExclamation, sMensaje
      Exit Sub
   End If

   Select Case Index
          Case Is = 0 ' Elimina
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsDetalle!lImprime) Then
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                  End If
                  Cn.Execute "update DGUIA SET tPedidoCodigo ='', tItemPedido ='', lFlagAgrega ='0' Where tPedidoCodigo ='" & sPedido & "' and tItemPedido ='" & sItem & "'"
                  If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And RsDetalle!lImprime) Then
                     wCabecera = False
                     ActivaFrame fraEliminacion
                     ActivaCabecera False
                     ActivaOpcion False
                     cmdOpcion(1).Enabled = False
                     cmdOpcion(2).Enabled = False
                  Else
                     sCodigo = ""
                     sDescrip = ""
                     EliminaItem
                  End If
              End If
                   
          Case Is = 1 ' Cantidad
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And Val(sDescrip) > 0 Then
                  nCantidad = Val(sDescrip)
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "update DPEDIDO set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                  RsDetalle.Requery
                  RsDetalle.AbsolutePosition = nPos
                  txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               End If
               
          Case Is = 2 ' Aumentar
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               nCantidad = nCantidad + 1
               Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
               txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               
          Case Is = 3 ' Disminuir
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdDetalle.Columns(3).Text > 1 Then
                  nCantidad = nCantidad - 1
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                  RsDetalle.Requery
                  RsDetalle.AbsolutePosition = nPos
                  txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               End If
                    
          Case Is = 4 ' Propiedad
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                         
               ActivaFrame fraPropiedad
               ActivaCabecera False
               ActivaOpcion False
               
               cmdOpcion(1).Enabled = False
               cmdOpcion(2).Enabled = False
               cmdOpcion(9).Enabled = True
               
          Case Is = 5 ' Observacion
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               frmKeyBoard.Caption = "Observación del Producto"
               frmKeyBoard.txtResultado.Text = lblObservacion.Caption
               frmKeyBoard.Show vbModal
               If wEnter = True Then
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "Update DPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "'"
                  lblObservacion.Caption = sDescrip
               End If
          
          Case Is = 6 ' Precios
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If Supervisor("03") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               wDetalle = True
               ActivaFrame fraDetalle
               ActivaCabecera False
               ActivaOpcion False
               cmdOpcion(1).Enabled = True
               cmdOpcion(2).Enabled = True
                              
   End Select
End Sub

Private Sub cmdEliminacion_Click(Index As Integer)
    
    RsMotivoEliminacion.MoveFirst
    RsMotivoEliminacion.Find ("Descripcion = '" & cmdEliminacion(Index).Caption & "'")
    
    If RsMotivoEliminacion.EOF Then
       RsMotivoEliminacion.MoveFirst
    End If
    
    If RsMotivoEliminacion!Codigo = "000" Then
       frmKeyBoard.txtResultado = ""
       frmKeyBoard.Show vbModal
       If Not wEnter Then
          Exit Sub
       End If
       sCodigo = "000"
    Else
       sCodigo = RsMotivoEliminacion!Codigo
       sDescrip = ""
    End If
    
    If wCabecera Then
       EliminaCabecera
    Else
       EliminaItem
    End If
        
    fraEliminacion.Visible = False
    wDetalle = True
    wAgrega = False
End Sub

Private Sub cmdGrupo_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 5
       cmdGrupo(i).BackColor = vbButtonFace
       cmdSubGrupo(i).BackColor = vbButtonFace
   Next i
   cmdSubGrupo(6).BackColor = vbButtonFace
   cmdSubGrupo(7).BackColor = vbButtonFace
   
   RsGrupo.MoveFirst
   RsGrupo.Find "nboton = " & Trim(Str(Index))
   sGrupo = RsGrupo!Codigo
   xGrupo = RsGrupo!Codigo
   cmdGrupo(RsGrupo!nBoton).BackColor = vbBlue
   RsSubGrupo.Filter = "tGrupo = '" & sGrupo & "'"
   AsignaBoton 7, RsSubGrupo, cmdSubGrupo()
   
   If RsSubGrupo.RecordCount = 0 Then
      sSubGrupo = ""
   Else
      RsSubGrupo.MoveFirst
      RsSubGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "' and lActivo=1", Cn) & "'"
      sSubGrupo = RsSubGrupo!Codigo
      xSubGrupo = RsSubGrupo!Codigo
      If RsSubGrupo!nBoton > 0 Then
         cmdSubGrupo(RsSubGrupo!nBoton).BackColor = vbRed
      End If
 
      RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido
   End If
   Screen.MousePointer = vbDefault
   txtBarra.SetFocus
End Sub

Private Sub cmdImpuesto_Click(Index As Integer)
    Select Case Index
        Case Is = 0
             nImpuesto1 = IIf(nImpuesto1 = 0, nPBase * nPorcentaje1 / 100, 0)
             txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
        Case Is = 1
             nImpuesto2 = IIf(nImpuesto2 = 0, nPBase * nPorcentaje2 / 100, 0)
             txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
        Case Is = 2
             nImpuesto3 = IIf(nImpuesto3 = 0, nPBase * nPorcentaje3 / 100, 0)
             txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    End Select
    nPVenta = nPBase + nImpuesto1 + nImpuesto2 + nImpuesto3
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    txtBarra.SetFocus
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdCabecera
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdCabecera
           Case Is = 2 'Previo
                MoverPuntero previo, grdCabecera
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdCabecera
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdCabecera
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdCabecera
           Case Is = 6 'Primero
                MoverPuntero Primero, grdDetalle
           Case Is = 7 'PgUp
                MoverPuntero pgup, grdDetalle
           Case Is = 8 'Previo
                MoverPuntero previo, grdDetalle
           Case Is = 9 'Siguiente
                MoverPuntero siguiente, grdDetalle
           Case Is = 10 'PgDn
                MoverPuntero pgdn, grdDetalle
           Case Is = 11 'Ultimo
                MoverPuntero Ultimo, grdDetalle
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   cmdOpcion(1).Enabled = True
   Dim i As Integer
   txtBarra.SetFocus
   Select Case Index
       Case Is = 0  'Exit
            Unload Me
            
       Case Is = 1  'Grabar
            If wDetalle Then
               ActivaOpcion True
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
               GrabaProducto
            Else
               If sCliente = "" And sTipoPedido = "02" Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               ActivaOpcion True
               ActivaFrame fraPlato
               ActivaCabecera True
                              
               If wAgrega Then
                  wAgrega = False
                   
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select nCORRELATIVO as Codigo from TPARAMETRO", Cn) + 1
                  If IsNull(nCorrela) Then
                      nCorrela = 1
                  End If
                  Cn.Execute "Update TPARAMETRO Set nCorrelativo = " & nCorrela

                  'Obtiene el Numero de Pedido
                  sPedido = Calcular("select max(tCodigoPedido) as Codigo from mpedido where substring(tCodigoPedido,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(sPedido) Or Mid(sPedido, 1, 2) <> Mid(Trim(Str(Year(Date))), 3, 2) Then
                      sPedido = Mid(Trim(Str(Year(Date))), 3, 2) & "00000001"
                  Else
                      sPedido = Mid(Trim(Str(Year(Date))), 3, 2) & Lib.Correlativo(Mid(sPedido, 3, 8), 8)
                  End If
                                                      
                  Isql = "insert into MPEDIDO " & _
                         "(tCodigoPedido,nCorrelativo, tClienteDelivery,fFecha,tEstadoPedido,tTipoPedido,lPrioridad, " & _
                         " tTipoAtencion,tMesa,tMozo,tMotorizado,tCaja, tSalon, tTurno, tObservacion,fProgramacion,nTiempo,tUsuario,nAdulto,nNino,nMesa,tPuntoVenta,tHabitacion, tReserva, tPasajero, tCompania, tContacto, tCajaAnterior, tTurnoAnterior, nDescuento, tDescuento, tObservacionDescuento, tUsuarioDescuento, fRegistro)" & _
                         "Values(   '" & sPedido & "'," _
                                       & nCorrela & "," _
                                 & "'" & sCliente & "'," _
                                 & "getdate() ," _
                                 & "'01'," _
                                 & "'" & sTipoPedido & "'," _
                                 & "'" & IIf(txtTipoPedido.ForeColor = &H800000, 0, 1) & "'," _
                                 & "'" & sTipoAtencion & "'," _
                                 & "''," _
                                 & "''," _
                                 & "''," _
                                 & "'" & sCaja & "'," _
                                 & "'" & sSalon & "'," _
                                 & "'" & sTurno & "'," _
                                 & "'" & txtObservacion.Caption & "'," _
                                 & "Null, " _
                                 & ntTiempo & ", '" & sUsuario & "', 0, 0, 0, " _
                                 & "'', '" & sHabitacion & "', '', '', '', '', " _
                                 & "'" & IIf(Not lAdicion, sCaja, "") & "', '" & IIf(Not lAdicion, sTurno, "") & "', " & xDescuento & ", '" & sCodigoDescuento & "', '" & IIf(sCodigoDescuento = "000", sDescripcionDescuento, "") & "', " _
                                 & "'" & tAutorizaDescuento & "', getDate())"
                  Cn.Execute Isql
                  RsCabecera.Requery
                  RsCabecera.MoveFirst
                  RsCabecera.Find "Codigo = '" & sPedido & "'"
               Else
                  Isql = "Update MPEDIDO Set " & _
                          "tClienteDelivery ='" & sCliente & "', " & _
                          "tTipoPedido ='" & sTipoPedido & "', " & _
                          "lPrioridad ='" & IIf(txtTipoPedido.ForeColor = &H800000, 0, 1) & "', " & _
                          "tTipoAtencion ='" & sTipoAtencion & "', " & _
                          "tObservacion ='" & txtObservacion.Caption & "', " & _
                          "nTiempo =" & ntTiempo & ", " & _
                          "nDescuento=" & xDescuento & ", " & _
                          "tDescuento='" & sCodigoDescuento & "', tUsuarioDescuento='" & tAutorizaDescuento & "', " & _
                          "tObservacionDescuento='" & IIf(sCodigoDescuento = "000", sDescripcionDescuento, "") & "' " & _
                          " where tCodigoPedido = '" & sPedido & "'"
                          
                    Cn.Execute Isql
                    RsCabecera.Requery
                    RsCabecera.Find "Codigo = '" & sPedido & "'"
               End If
            End If
            
            If wCalcula Then
               CalculaDescuento
               RsDetalle.Requery
               wCalcula = False
            End If
            
            ActivaOpcion True
            grdCabecera.Enabled = True
            
       Case Is = 2  'Cancelar
            ActivaOpcion True
            If wDetalle Then
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
            Else
               If RsCabecera.RecordCount = 0 Then
                  ActivaCabecera False
                  cmdCabecera(0).Visible = True
               Else
                  ActivaCabecera True
                  ActivaFrame fraPlato
               End If
               AsignaPedido
               wAgrega = False
            End If
            grdCabecera.Enabled = True
       
       Case Is = 5  'Pre cuenta
            If RsCabecera.RecordCount = 0 Then
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
            
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
               If i > 0 Then
                  MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            wEnter = False
            Screen.MousePointer = vbHourglass
            Cn.Execute "update MPEDIDO set nPrecuenta = isnull(nPrecuenta,0) + 1   where tCodigoPedido='" & sPedido & "'"
            If sTipoPedido = "02" Then
               Isql = "select * from vPreCuentaDelivery WHERE Codigo='" & sPedido & "'"
            Else
               If lPrecuenta Then
                  Isql = "select * from vPrecuentaDetallada WHERE Codigo='" & sPedido & "'"
               Else
                  If lPrecuentaAgrupada Then
                     Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                  Else
                     Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                  End If
               End If
            End If

            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               If sTipoPedido = "02" Then
                  ImprimeDelivery RsImpresion, RsProductoPropiedad
               Else
                  If lPrecuenta Then
                     ImprimePreCuentaDetallada RsImpresion, sCodigo
                  Else
                     If lInfhotel Then
                        ImprimeInfhotel RsImpresion, sCodigo
                     Else
                        ImprimePreCuenta RsImpresion, sCodigo
                     End If
                  End If
               End If
               Cn.Execute "update MPEDIDO set fRegCuenta = getDate()  where tCodigoPedido='" & sPedido & "'"
            End If
            LimpiaRs
                                
       Case Is = 6  'Pago
            If RsTipoDocumento.RecordCount = 0 Then
               Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            sVar1 = sTipoPedido
            frmDocumento.Show vbModal
            RsCabecera.Requery
            RsDetalle.Requery
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If
                                                
       Case Is = 7  'Código Directo
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            sTemp = ""
            Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
            Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                            "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                            "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                            "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                            "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
            frmBusquedaRapida.nPredeterm = 1
            frmBusquedaRapida.Show vbModal
   
            If wEnter Then
               sProducto = sCodigo
               Dim xxx As String
               xxx = RsProducto.Filter
               RsProducto.Filter = adFilterNone
               RsProducto.MoveFirst
               RsProducto.Find "Codigo = '" & sProducto & "'"
               
               If vComanda Then
                  sTipo = "Comanda"
                  sDescrip = sComanda
                  frmNumPad.Caption = "Comanda"
                  frmNumPad.Show vbModal
                
                  sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                  
                  If wComanda And sComanda = "" Then
                     MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                     Exit Sub
                  End If
               
                  If Not wEnter Then
                     Exit Sub
                  End If
               End If
               
               InsertaProducto 1
               RsProducto.Filter = xxx
                
               If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                  lPropiedad = True
               End If
            
            End If
            txtBarra.SetFocus
                        
       Case Is = 8  'Importar Pedido
            sTemp = ""
            Isql = "SELECT dbo.MGUIA.tCodigoPedido AS Codigo, dbo.DGUIA.tItem as Descripcion, dbo.MGUIA.fEntrega, dbo.DGUIA.nCantidad, dbo.DGUIA.tCodigoProducto, dbo.TPRODUCTO.tResumido AS Producto, dbo.TMODELO.tResumido AS Modelo " & _
                   "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DGUIA LEFT OUTER JOIN dbo.MGUIA ON dbo.DGUIA.tCodigoPedido = dbo.MGUIA.tCodigoPedido LEFT OUTER JOIN dbo.TMODELO ON dbo.DGUIA.tCodigoModelo = dbo.TMODELO.tCodigoModelo ON dbo.TPRODUCTO.tCodigoProducto = dbo.DGUIA.tCodigoProducto " & _
                   "where tClienteCorp='" & sCliente & "' and tTienda='" & sTienda & "' and lFlagAgrega = 0"
                   
            Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Pedido", 2, "Codigo", 1100, 2, 0, "", _
                                                            "Entrega", 2, "fEntrega", 1000, 2, 0, "dd/mm/yyyy", _
                                                            "Item", 2, "Descripcion", 600, 2, 0, "", _
                                                            "Cant", 2, "nCantidad", 800, 1, 0, "##,##0.00", _
                                                            "Producto", 2, "Producto", 2700, 0, 0, "", _
                                                            "Modelo", 2, "Modelo", 2000, 0, 0, "")
            frmBusquedaRapida.nPredeterm = 0
            frmBusquedaRapida.Show vbModal
   
            If wEnter Then
               Dim RsGuia As Recordset
               Set RsGuia = Lib.OpenRecordset("select tCodigoProducto, nCantidad from DGUIA where tCodigoPedido='" & sCodigo & "' and tItem='" & sDescrip & "'", Cn)
               If RsGuia.RecordCount = 0 Then
                  MsgBox "Error Fatal", vbCritical, sMensaje
                  Exit Sub
               End If
                              
               Dim sFiltro As String
               sProducto = RsGuia!tCodigoProducto
               sFiltro = RsProducto.Filter
               RsProducto.Filter = adFilterNone
               RsProducto.MoveFirst
               RsProducto.Find "Codigo = '" & sProducto & "'"
               If RsProducto.EOF Then
                  MsgBox "Error Fatal, Producto no encontrado", vbCritical, sMensaje
                  Exit Sub
               End If
               
               InsertaProducto RsGuia!nCantidad
               RsProducto.Filter = sFiltro
                 
               Isql = "UPDATE DGUIA SET " _
                    & " tPedidoCodigo ='" & sPedido & "', " _
                    & " tItemPedido ='" & sItem & "', " _
                    & " lFlagAgrega ='1'" _
                    & " Where tCodigoPedido ='" & sCodigo & "' AND tItem ='" & sDescrip & "'"
               Cn.Execute Isql
            End If
            
            
            
       Case Is = 9  'Salir Propiedades
            RsDetalle.Requery
            RsDetalle.MoveFirst
            RsDetalle.Find "titem = '" & sItem & "'"
            ActivaOpcion True
            ActivaCabecera True
            ActivaFrame fraPlato
                    
       Case Is = 10  'Busca Pedido
            If Not RsCabecera.EOF Then
               wEnter = False
               Sw = False
               sDescrip = ""
               frmIr.Show vbModal
               
               If wEnter Then
                  RsCabecera.MoveFirst
                  RsCabecera.Find ("tMesa='" & sCodigo & "'")
                  If RsCabecera.EOF Then
                     RsCabecera.MoveFirst
                     MsgBox "El pedido pertenece a otra caja ", vbExclamation, sMensaje
                  End If
                  grdCabecera.Bookmark = RsCabecera.Bookmark
               End If
               
               If Sw Then
                  RsCabecera.MoveFirst
                  If Len(Trim(sDescrip)) > 0 Then
                     RsCabecera.Find ("tObservacion like '*" & Trim(sDescrip) & "*'")
                     If RsCabecera.EOF = True Then
                        RsCabecera.MoveFirst
                        MsgBox "No se encontro " & sDescrip, vbExclamation, sMensaje
                        Exit Sub
                     End If
                     grdCabecera.Bookmark = RsCabecera.Bookmark
                  End If
                  Sw = False
               End If
            End If
            
       Case Is = 11 ' Refrescar
            RsCabecera.Requery
            RsDetalle.Requery
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If
                   
       Case Is = 13 ' Visualizar Pedido
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            sTipo = ""
            frmPedido.Show vbModal
            
       Case Is = 14 ' Cuentas internas
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
       
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
                        
            sUsuarioAutoriza = sUsuario
            If Supervisor("09") = False Then
               MsgBox "Clave no permitida", vbExclamation, sMensaje
               Exit Sub
            End If
            sUsuarioAutoriza = sVar1
                        
            If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cuenta Corrientes, pedidos con items Facturados", vbExclamation, sMensaje
               Exit Sub
            End If
                              
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If
                                                                        
            sTemp = ""
            Isql = "select * from vCompania where lActivo = 1 order by Descripcion"
            Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 800, 2, 0, "", _
                                                            "Identidad", 2, "Identidad", 1500, 0, 0, "", _
                                                            "Cliente", 2, "Descripcion", 2800, 0, 0, "", _
                                                            "Linea", 2, "nLinea", 1050, 1, 0, "##,##0.00", _
                                                            "Consumo", 2, "nConsumo", 1050, 1, 0, "##,##0.00", _
                                                            "Fecha Venc", 2, "fFechaVence", 1050, 0, 0, "dd/mm/yyyy")
            frmBusquedaRapida.nPredeterm = 2
            frmBusquedaRapida.Show vbModal
                                                                         
            If Not wEnter Or sCodigo = "" Then
               Exit Sub
            End If

            sCliente = sCodigo
            xsuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido='" & sPedido & "'", Cn)
            xLinea = Calcular("select nLinea as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
            xConsumo = Calcular("select nConsumo as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
            
            If xsuma > xLinea - xConsumo Then
               MsgBox "El Cliente no tiene linea suficiente " & Chr(13) & _
                      "Linea : " & Format(xLinea, "###,##0.00") & "  Consumo : " & Format(xConsumo, "###,##0.00") & Chr(13) & _
                      "Saldo : " & Format(xLinea - xConsumo, "###,##0.00"), vbCritical, sMensaje
               Exit Sub
            End If

            If MsgBox("Esta seguro de Enviar el Pedido Nro: " & sPedido & _
               Chr(13) & "a Cuentas Corrientes al Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               Exit Sub
            End If
                        
            Cn.Execute "Update MPEDIDO set tClienteCtaCte ='" & sCliente & "', tEstadoPedido = '04'  where tCodigoPedido='" & sPedido & "'"
            Cn.Execute "Update TCOMPANIA set nConsumo = " & xConsumo + xsuma & " where tCodigoCliente ='" & sCliente & "'"
            
            Isql = "select * from vCtaCte " & _
                   "WHERE Codigo='" & sPedido & "'"
            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               ImprimeCtaCte RsImpresion
            End If
            LimpiaRs
                        
            RsCabecera.Requery
            RsDetalle.Requery
            
            If RsCabecera.RecordCount = 0 Then
               sPedido = ""
               ActivaCabecera False
               cmdCabecera(0).Visible = True
               fraPlato.Visible = False
            End If
                                                
       Case Is = 16 ' Caja Rapida
            Screen.MousePointer = vbHourglass
            frmCajaRapida.Show vbModal
            RsCabecera.Requery
            RsDetalle.Requery
                                        
       Case Is = 17 ' Cancelar del Motivo de Eliminacion
            ActivaFrame fraPlato
            ActivaCabecera True
            ActivaOpcion True
            
   End Select
End Sub

Private Sub cmdOperador_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 13
       cmdOperador(i).BackColor = vbButtonFace
   Next i
   RsOperador.MoveFirst
   RsOperador.Find "nboton = " & Trim(Str(Index))
   cmdOperador(Index).BackColor = vbRed
   AsignaPropiedad
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrecio_Click()
   If nDecimal > 2 Then
      sTipo = "Decimal4"
   Else
      sTipo = ""
   End If
   frmNumPad.Show vbModal
   If wEnter Then
      nPVenta = Val(sDescrip)
      nOficial = nPVenta
      nDescuento = 0
      nRecargo = 0
      txtDPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtRPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      Acumulado = 0
      Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
      Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
      Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
      Acumulado = 1 + (Acumulado / 100)
      nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
      nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
      nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
      
      nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtOficial.Caption = Format(nOficial, "###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
      txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
   End If
End Sub

Private Sub cmdProducto_Click(Index As Integer)
    txtBarra.SetFocus

    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
        
    If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tCaja='" & sCaja & "' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
       MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
       RsCabecera.Requery
       Exit Sub
    End If
    
    RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Find "nboton = " & Trim(Str(Index))
    sProducto = RsProducto!Codigo
    
    If vComanda Then
       sTipo = "Comanda"
       sDescrip = sComanda
       frmNumPad.Caption = "Comanda"
       frmNumPad.Show vbModal
       
       sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
       
      If wComanda And sComanda = "" Then
         MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
         Exit Sub
      End If
      
      If Not wEnter Then
          Exit Sub
       End If
    End If
        
    InsertaProducto 1
    
    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
       lPropiedad = True
    End If
    
End Sub

Private Sub cmdPropiedad_Click(Index As Integer)
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & cmdPropiedad(Index).Caption & "'")
    
    If cmdPropiedad(Index).FontBold = True Then
       cmdPropiedad(Index).FontBold = False
       If Not RsPropiedad.EOF Then
          Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!Codigo & "'"
          If RsPropiedad!nPrecio > 0 Then
             CambiaPrecio nPVenta - RsPropiedad!nPrecio
          End If
          lblResumen.Caption = Replace(lblResumen.Caption, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
       End If
    Else
       cmdPropiedad(Index).FontBold = True
       If Not RsPropiedad.EOF Then
          Cn.Execute "Insert into TPRODUCTOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & RsPropiedad!Codigo & "', '" & sProducto & "', '" & RsPropiedad!tInsumo & "', " & IIf(IsNull(RsPropiedad!nCantidad), 0, RsPropiedad!nCantidad) & ", 0)"
          If RsPropiedad!nPrecio > 0 Then
             CambiaPrecio nPVenta + RsPropiedad!nPrecio
          End If
          lblResumen.Caption = lblResumen.Caption & RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", "
       End If
    End If
    RsProductoPropiedad.Requery

End Sub

Private Sub cmdSubGrupo_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 7
       cmdSubGrupo(i).BackColor = vbButtonFace
   Next i
   
   RsSubGrupo.MoveFirst
   RsSubGrupo.Find "nboton = " & Trim(Str(Index))
   xSubGrupo = RsSubGrupo!Codigo
   cmdSubGrupo(RsSubGrupo!nBoton).BackColor = vbRed
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido
   Screen.MousePointer = vbDefault
   txtBarra.SetFocus
End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim nTotal As Double
    
    sUsuarioAutoriza = sUsuario
    
    wConsumo = False
    sVar2 = sTipoPedido
    txtBarra.SetFocus
    If RsDetalle.RecordCount = 0 Then
       Exit Sub
    End If

    'Chequea Consistencia Infhotel
    
    If lPrinter And lObligaPrinter Then
       i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
       If i > 0 Then
          MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
          Exit Sub
       End If
    End If

    If lObligaPrecuenta Then
       i = Calcular("select nPrecuenta as codigo from MPEDIDO where tCodigoPedido ='" & sPedido & "'", Cn)
       If i = 0 Then
          MsgBox "Falta Imprimir la Precuenta", vbExclamation, sMensaje
          Exit Sub
       End If
    End If

    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    End If
            
'    'Chequea Descuento
'    If CDbl(txtDescuento) > 0 Then
'       Dim nConsumo As Double
'       Dim nActual As Double
'       Dim nTope As Double
'       Dim sTexto As String
'
'       If Calcular("select lPrecioVenta as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) = False Then
'          sTexto = "DPEDIDO.nPrecioNeto*DPEDIDO.nCantidad"
'       Else
'          sTexto = "DPEDIDO.nVenta"
'       End If
'
'       nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
'       Isql = "select sum(" & sTexto & ") as codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'              "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'"
'       nActual = Calcular("select sum(" & sTexto & ") as codigo From dbo.DPEDIDO WHERE tCodigoPedido='" & sPedido & "' and tEstadoItem='N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0)", Cn)
'       nConsumo = Calcular(Isql, Cn)
'
'       If nConsumo + nActual > nTope And nTope > 0 Then
'          MsgBox "El Consumo exede al Tope" & Chr(13) & _
'                 "Tope : " & Format(nTope, "###,##0.00") & Chr(13) & _
'                 "Saldo: " & Format(nTope - nConsumo, "###,##0.00") & Chr(13), vbCritical, sMensaje
'            Exit Sub
'       End If
'    End If
            
    'Verifica si existe item por facturar
    X = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
    If X <= 0 Then
       MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
       Exit Sub
    End If

    'Por Consumo
    If RsTipoDocumento!tTipoEmision <> "00" Then
       If lConsumo1 = True Then
          nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
          If nRespuesta = vbYes Then
             frmKeyBoard.txtResultado = "POR CONSUMO"
             frmKeyBoard.Show vbModal
             If sDescrip = "" Or Not wEnter Then
                MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                Exit Sub
             End If
             sdetalleconsumo = sDescrip
             wConsumo = True
          ElseIf nRespuesta = vbCancel Then
             Exit Sub
          End If
        End If
    End If

    'Verifica si no ha sido importado
    i = Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tCodigoPedido ='" & sPedido & "' and tCaja ='" & sCaja & "'", Cn)
    If i <= 0 Then
       MsgBox "Error: Pedido jalado por otra caja, favor de Refrescar", vbCritical, sMensaje
       Exit Sub
    End If

    'Tipo de Emision
    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
       'Factura
       'Consistencia Factura

       'Calcula el total de la cabecera
       nTotal = Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
       
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
          sTemp = ""
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          frmBusquedaRapida.nPredeterm = 1
          Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
          If sTipoPedido = "02" Then
             sTemp = Calcular("SELECT dbo.TCLIENTE.tIdentidad AS codigo FROM dbo.TDELIVERY INNER JOIN dbo.TCLIENTE ON dbo.TDELIVERY.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente WHERE tCodigoDelivery ='" & sCliente & "'", Cn)
             sTemp = IIf(sTemp = "0", "", sTemp)
          Else
             sTemp = ""
          End If
          frmBusquedaRapida.Show vbModal
          If wEnter = True And sCodigo <> "" Then
             If sTipoPedido = "02" Then
                Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sCodigo & "' where tCodigoDelivery ='" & sCliente & "'"
             End If
             sTemp = sCliente
             sCliente = sCodigo
          Else
             Exit Sub
          End If
       End If
              
       'Genera y Actualiza los Numero de Documento
       RsDetalle.MoveFirst

       For i = 1 To IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
           RsTipoDocumento.Requery
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
           If RsTipoDocumento.EOF Then
              MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
              Exit Sub
           End If

           sSerie = RsTipoDocumento!tSerie
           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
           sPrefijo = RsTipoDocumento!prefijo
           sTipoDocumento = RsTipoDocumento!tTipoEmision
           sImp = RsTipoDocumento!tImpresora
           sDocumento = sPrefijo & sSerie & sCorrela
           sResumen = RsTipoDocumento!lResumen
           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

           'Genera el Detalle de DDOCUMENTO
           Dim xClave As String
           For j = 1 To nFItem
               xClave = RsDetalle!tItem
               Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) "
               Cn.Execute Isql
               RsDetalle.MoveFirst
               RsDetalle.Find ("tItem ='" & xClave & "'")
               RsDetalle.MoveNext
               If RsDetalle.EOF Then
                  Exit For
               End If
           Next j

           'Inserta Detalle de Documento
           Isql = "Insert into DDOCUMENTO " & _
                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                  "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
                  "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tDocumento ='" & sDocumento & "'"
           Cn.Execute Isql

           'Genera el Detalle de MDOCUMENTO
           Isql = "Insert into MDOCUMENTO " & _
                  "     ( tDocumento, tTipoDocumento, tCodigoCliente, tEstadoDocumento, tCaja, tTurno, tSalon, tUsuario, tUsuarioAutoriza, fRegistro) " & _
                  "Values(   '" & sDocumento & "', " _
                          & "'" & sTipoDocumento & "', " _
                          & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                          & "'01', " _
                          & "'" & sCaja & "', " _
                          & "'" & sTurno & "', " _
                          & "'" & sSalon & "', " _
                          & "'" & sUsuario & "', " _
                          & "'" & sUsuarioAutoriza & "', " _
                          & " getdate()) "
           Cn.Execute Isql

           'Calcula el total de la cabecera
           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                          " from DPEDIDO where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)

           'Actualiza el Documento con el Temporal
           nCargo = Round(RsSuma!nVenta, 2)
           Isql = "Update MDOCUMENTO set nNeto = " & RsSuma!nNeto & " , " & _
                                        "nRecargo = 0, " & _
                                        "nDescuento = 0, " & _
                                        "nPrecioOficial = 0 , " & _
                                        "nPrecioImpuesto1 = " & RsSuma!nImpuesto1 & " , " & _
                                        "nPrecioImpuesto2 = " & RsSuma!nImpuesto2 & " , " & _
                                        "nPrecioImpuesto3 = " & RsSuma!nImpuesto3 & " , " & _
                                        "nVenta = " & RsSuma!nVenta & _
                                        ",lreplica=1 where tDocumento = '" & sDocumento & "'"
           Cn.Execute Isql

           'Actualiza Base de Datos Detalle del Pedido
           Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
           Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & sPedido & "'"
           
           'Imprime Documentos
           Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.fRegistro, dbo.MPEDIDO.tObservacion, " & _
                  "dbo.MPEDIDO.tCodigoPedido, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tCodigoProducto, dbo.TPRODUCTO.tDetallado AS Producto, MIN(dbo.DDOCUMENTO.tItem) AS tItem, SUM(dbo.DDOCUMENTO.nCantidad) AS nCantidad, SUM(dbo.DDOCUMENTO.nVenta) AS nPrecioVenta " & _
                  "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN dbo.MPEDIDO ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
                  "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MPEDIDO.tCodigoPedido, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.TCLIENTE.tEmpresa, dbo.TCLIENTE.tIdentidad, dbo.TCLIENTE.tDireccion, " & _
                  "dbo.MPEDIDO.tObservacion, dbo.DDOCUMENTO.tCodigoProducto, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.TPRODUCTO.tDetallado " & _
                  "HAVING (dbo.MDOCUMENTO.tDocumento = N'" & sDocumento & "') order by tItem"
           
           Set RsImpresion = Lib.OpenRecordset(Isql, Cn)

           If RsImpresion.RecordCount = 0 Then
              LimpiaRs
              MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
           Else
              'Configura la Impresora
              Imprimir (sImp)
              Printer.FontName = sFont
              Printer.FontBold = False

              If sTipoDocumento = "01" Then
                 If wConsumo Then
                    ImprimeFacturaConsumoN RsImpresion, sdetalleconsumo, sEmpresa
                 Else
                    ImprimeFacturaN RsImpresion, sEmpresa
                 End If
                 nFactura = sCorrela
                 lblFactura.Caption = nFactura
              Else
                 If wConsumo Then
                    ImprimeBoletaConsumoN RsImpresion, sdetalleconsumo, sEmpresa
                 Else
                    ImprimeBoletaN RsImpresion, sEmpresa
                 End If
              End If
           End If
       Next i

       If i - 1 = 1 Then
          LimpiaRs
          'Cancelacion del Documento

          xTipo = ""
          If sVar2 = "02" Then
             frmPrePago.Show vbModal
             If wEnter Then
                cmdOpcion_Click (5)
             End If
          Else
             sFormulario = ""
             sTipo = ""
             frmPago.Show vbModal
          End If
            
          If wEnter Then
             'Actualiza Base de Datos Cabecera del Pedido
             Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & sPedido & "'"
          End If
          
          RsCabecera.Requery
          RsDetalle.Requery
       Else
          frmDocumento.Show vbModal
       End If

    Else
       'Calcula el total de la cabecera
       nTotal = Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)

       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
          'Factura
          'Consistencia Factura
          frmBusquedaRapida.nPredeterm = 1
          If sTipoPedido = "02" Then
             sTemp = Calcular("SELECT dbo.TCLIENTE.tIdentidad AS codigo FROM dbo.TDELIVERY INNER JOIN dbo.TCLIENTE ON dbo.TDELIVERY.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente where tCodigoDelivery ='" & sCliente & "'", Cn)
             sTemp = IIf(sTemp = "0", "", sTemp)
          Else
             sTemp = ""
          End If
                    
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
          Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          frmBusquedaRapida.Show vbModal
          
          If wEnter = True And sCodigo <> "" Then
             sTemp = sCliente
             sCliente = sCodigo
          Else
             Exit Sub
          End If
       Else
          sCliente = ""
       End If

       'Otros
       'Genera y Actualiza los Numero de Documento
       sSerie = RsTipoDocumento!tSerie
       sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
       sPrefijo = RsTipoDocumento!prefijo
       sTipoDocumento = RsTipoDocumento!tTipoEmision
       sImp = RsTipoDocumento!tImpresora
       sDocumento = sPrefijo & sSerie & sCorrela
       sResumen = RsTipoDocumento!lResumen
       Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

       'Inserta Detalle de Documento
       Isql = "Insert into DDOCUMENTO " & _
              "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
              "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
              "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
              "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
              "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
       Cn.Execute Isql

       'Calcula el total de la cabecera
       Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                      " from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "' group by tCodigoPedido ", Cn)

       'Inserta el Documento
       nCargo = Round(RsSuma!nVenta, 2)
       Isql = "Insert into MDOCUMENTO " & _
              "     ( tDocumento, tTipoDocumento, tCortesia, tcodigoCliente, tEstadoDocumento, tCaja, tTurno, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, tSalon, tUsuario, tUsuarioAutoriza, fRegistro) " & _
              "Values(   '" & sDocumento & "', " _
                      & "'" & sTipoDocumento & "', " _
                      & "'" & sCortesia & "', " _
                      & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                      & "'01', " _
                      & "'" & sCaja & "', " _
                      & "'" & sTurno & "', " _
                      & RsSuma!nNeto & ", " _
                      & "0, 0, 0, " _
                      & RsSuma!nImpuesto1 & ", " _
                      & RsSuma!nImpuesto2 & ", " _
                      & RsSuma!nImpuesto3 & ", " _
                      & RsSuma!nVenta & ", " _
                      & "'" & sSalon & "', " _
                      & "'" & sUsuario & "', " _
                      & "'" & sUsuarioAutoriza & "', " _
                      & "getdate() ) "
       Cn.Execute Isql

       'Actualiza Base de Datos Detalle del Pedido
        If sPrefijo = "0" Then
           Cn.Execute "Update DPEDIDO set tFacturado = 'C', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tfacturado))=0)"
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido = '" & sPedido & "'"
           Cn.Execute "Update MDOCUMENTO set tEstadoDocumento ='02',lreplica=1 where tDocumento = '" & sDocumento & "'"
        Else
           Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
           Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & sPedido & "'"
        End If
                                         
       'Imprime Documentos
       'Isql = "SELECT * from vDOCUMENTOImpresora where tDocumento ='" & sDocumento & "' order by tItem"
           Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.fRegistro, dbo.MPEDIDO.tObservacion, " & _
                  "dbo.MPEDIDO.tCodigoPedido, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tCodigoProducto, dbo.TPRODUCTO.tDetallado AS Producto, MIN(dbo.DDOCUMENTO.tItem) AS tItem, SUM(dbo.DDOCUMENTO.nCantidad) AS nCantidad, SUM(dbo.DDOCUMENTO.nVenta) AS nPrecioVenta " & _
                  "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN dbo.MPEDIDO ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
                  "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MPEDIDO.tCodigoPedido, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.TCLIENTE.tEmpresa, dbo.TCLIENTE.tIdentidad, dbo.TCLIENTE.tDireccion, " & _
                  "dbo.MPEDIDO.tObservacion, dbo.DDOCUMENTO.tCodigoProducto, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.TPRODUCTO.tDetallado " & _
                  "HAVING (dbo.MDOCUMENTO.tDocumento = N'" & sDocumento & "') order by tItem"
       
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)

       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
          'Configura la Impresora
          Imprimir (sImp)
          Printer.FontName = sFont
          Printer.FontBold = False

          If wConsumo Then
             If RsTipoDocumento!tFormulario = "01" Then
                If RsTipoDocumento!Cliente Then
                   ImprimeFacturaConsumoT RsImpresion, sdetalleconsumo
                   nFactura = sCorrela
                   lblFactura.Caption = nFactura
                Else
                   ImprimeBoletaConsumoT RsImpresion, sdetalleconsumo
                End If
             Else
                If RsTipoDocumento!Cliente Then
                   ImprimeFacturaConsumoN RsImpresion, sdetalleconsumo, sEmpresa
                   nFactura = sCorrela
                   lblFactura.Caption = nFactura
                Else
                   ImprimeBoletaConsumoN RsImpresion, sdetalleconsumo, sEmpresa
                End If
             End If
          Else
             If RsTipoDocumento!tFormulario = "01" Then
                If RsTipoDocumento!Cliente Then
                   ImprimeFacturaT RsImpresion
                   nFactura = sCorrela
                   lblFactura.Caption = nFactura
                ElseIf RsTipoDocumento!tTipoEmision = "00" Then
                   If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                      ImprimeCortesia RsImpresion
                   End If
                Else
                   ImprimeBoletaT RsImpresion
                End If
             Else
                If RsTipoDocumento!Cliente Then
                   ImprimeFacturaN RsImpresion, sEmpresa
                   nFactura = sCorrela
                   lblFactura.Caption = nFactura
                ElseIf RsTipoDocumento!tTipoEmision = "00" Then
                   If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                      ImprimeCortesia RsImpresion
                   End If
                Else
                   ImprimeBoletaN RsImpresion, sEmpresa
                End If
             End If
          End If
       End If

       LimpiaRs

       'Cancelacion del Documento
       If RsTipoDocumento!tTipoEmision <> "00" Then
          Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & sPedido & "'"
          xTipo = ""
          If sVar2 = "02" Then
             frmPrePago.Show vbModal
             If wEnter Then
                cmdOpcion_Click (5)
             End If
          Else
             sFormulario = ""
             xTipo = ""
             frmPago.Show vbModal
          End If
       Else
          Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & sPedido & "'"
       End If
    End If
    
    If sTipoPedido = "02" Then
       Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sCliente & "' where tCodigoDelivery ='" & sTemp & "'"
    End If
    
    RsCabecera.Requery
    RsDetalle.Requery
    If RsCabecera.EOF Then
       LimpiaCabecera
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   If lRapido = False Then
      cmdOpcion(16).Visible = False
   End If
   
   frmVenta.Caption = " Caja : " & sCaja
   sComanda = ""
   sUsuarioAutoriza = ""
  
   'Cabecera
   Isql = "SELECT * from vPedidoCabecera " & _
          "where tEstadoPedido = '01' and tTipoPedido = '04' and tCaja = '" & sCaja & "' order by Codigo"
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
      
   'Detalle
   Isql = "select * from vPedidoDetalle " & _
          "where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tCaja = '" & sCaja & "') " & _
          "ORDER BY tCodigoPedido, tItem"
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
      
   'Grupo
   Isql = "select * from vGrupo where lActivo = 1 order by nBoton"
   Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
   If RsGrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un Grupo creado", vbCritical, sMensaje
      End
   End If
   
   'Sub Grupo
   Isql = "select * from vSubGrupo where lActivo = 1 Order by nBoton"
   Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
   If RsSubGrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un SubGrupo creado", vbCritical, sMensaje
      End
   End If
      
   'Producto
   Isql = "select * from vProducto where lActivo = 1 Order by nBoton"
   Set RsProducto = Lib.OpenRecordset(Isql, Cn)
   Set RsProducto = Lib.OpenRecordset(Isql, Cn)
   If RsProducto.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un producto creado", vbCritical, sMensaje
      End
   End If
                                                            
   'Motivo de Eliminacion
   Isql = "select * from vMotivoEliminacion where lActivo = 1 order by Codigo"
   Set RsMotivoEliminacion = Lib.OpenRecordset(Isql, Cn)
   AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()
            
   'Operador
   Isql = "select * from vOperador where lActivo = 1 order by Codigo"
   Set RsOperador = Lib.OpenRecordset(Isql, Cn)
   
   'Propiedades
   Isql = "select tCodigoPropiedad as Codigo, tDetallado as Descripcion, tProducto, tOperador, nPrecio, tInsumo, nCantidad from TPROPIEDAD where lActivo = 1"
   Set RsPropiedad = Lib.OpenRecordset(Isql, Cn)
         
   'Producto Propiedad
   Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, dbo.TPROPIEDAD.tDetallado AS Descripcion, dbo.vOperador.Descripcion AS Operador " & _
          "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto LEFT OUTER JOIN dbo.vOperador ON dbo.TPROPIEDAD.tOperador = dbo.vOperador.Codigo " & _
          "where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tCaja = '" & sCaja & "')"
   Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
      
   'Areas
   Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                                                                                                           
   'Tipo de Documentos
   Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and tTipoEmision <>'00' order by tTipoEmision", Cn)
   
   'Configuración de Grillas
   Call ConfGrilla(2, grdCabecera, "Cliente", 2, "Corp", 1200, 0, 0, "", _
                                   "Guia", 2, "tObservacion", 1600, 0, 0, "")
                                   
   Call ConfGrilla(8, grdDetalle, "Producto", 2, "Producto", 1990, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 750, 1, 0, "###,###,##0.00", _
                                  "Cant.", 2, "nCantidad", 700, 1, 0, "###0.00", _
                                  "SubTotal", 2, "nVenta", 980, 1, 0, "###,###,##0.00", _
                                  "F", 2, "tFacturado", 250, 0, 0, "", _
                                  "E", 2, "lImprime", 250, 0, 4, "", _
                                  "P", 2, "lPropiedad", 250, 0, 4, "", _
                                  "O", 2, "lObservacion", 250, 0, 4, "")
                                        
   Set grdCabecera.DataSource = RsCabecera
   Set grdDetalle.DataSource = RsDetalle
      
   'Asignar Grupo
   AsignaBoton 5, RsGrupo, cmdGrupo()
   RsGrupo.MoveFirst
   RsGrupo.Find "Codigo='" & sGrupoDefault & "'"
   If RsGrupo.EOF Then
      MsgBox "Error: Se necesita configurar en caja, un Grupo por defecto", vbCritical, sMensaje
      End
   End If
   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
      cmdGrupo(RsGrupo!nBoton).BackColor = vbBlue
   Else
      RsGrupo.MoveFirst
      RsGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vGrupo where nBoton>0 and lActivo=1", Cn) & "'"
      cmdGrupo(RsGrupo!nBoton).BackColor = vbBlue
   End If
   xGrupo = RsGrupo!Codigo
   RsSubGrupo.Filter = "tGrupo = '" & xGrupo & "'"
         
   'Asignar SubGrupo
   AsignaBoton 7, RsSubGrupo, cmdSubGrupo()
   RsSubGrupo.MoveFirst
   RsSubGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & xGrupo & "' and lActivo=1", Cn) & "'"
   If RsSubGrupo.EOF Or RsSubGrupo!nBoton = 0 Then
      MsgBox "Error: Se necesita configurar un SubGrupo con Boton", vbCritical, sMensaje
      End
   End If
   cmdSubGrupo(RsSubGrupo!nBoton).BackColor = vbRed
   xSubGrupo = RsSubGrupo!Codigo
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      
   'Asignar Producto
   sTipoPedido = "01"
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido
   
   'Asigna Operador
   AsignaBoton 13, RsOperador, cmdOperador()
   If RsOperador.RecordCount > 0 Then
      RsOperador.MoveFirst
      If Not IsNull(RsOperador!nBoton) And RsOperador!nBoton > 0 Then
         cmdOperador_Click (RsOperador!nBoton)
      End If
   End If
         
   AsignaComando 4, RsTipoDocumento, cmdTipoDocumento()
       
   ActivaFrame fraPlato
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!Codigo
   End If
   
   wDetalle = True
   wAgrega = False
    
   'Visualiza los Impuestos
   Impuesto
   lblFecha.Caption = Format(Date, "dddd, dd MMMM yyyy")
   lblFactura.Caption = nFactura
     
   lPropiedad = False
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaPedido()
   Screen.MousePointer = vbHourglass
   
   'Llena los valores
   If RsCabecera.RecordCount = 0 Then
      sTipoPedido = "04"
      sPedido = ""
      LimpiaCabecera
   Else
      sPedido = IIf(IsNull(RsCabecera!Codigo), "", RsCabecera!Codigo)
      grdDetalle.Caption = "Pedidos de " & grdCabecera.Columns(1).Text
      sCliente = IIf(IsNull(RsCabecera!tClienteCorp), "", RsCabecera!tClienteCorp)
      sTienda = IIf(IsNull(RsCabecera!tTienda), "", RsCabecera!tTienda)
      sTipoAtencion = IIf(IsNull(RsCabecera!tTipoAtencion), "01", RsCabecera!tTipoAtencion)
      sTipoPedido = IIf(IsNull(RsCabecera!tTipoPedido), "01", RsCabecera!tTipoPedido)
      nCorrela = IIf(IsNull(RsCabecera!nCorrelativo), "01", RsCabecera!nCorrelativo)
         
      fraPedido.Caption = " Pedido " & IIf(RsCabecera.RecordCount = 0, "", IIf(IsNull(RsCabecera!Codigo), "", RsCabecera!Codigo)) & " "
      txtFecha.Caption = IIf(IsNull(RsCabecera!fFecha), "", RsCabecera!fFecha)
      txtTipoPedido.Caption = IIf(IsNull(RsCabecera!TipoPedido), "", RsCabecera!TipoPedido)
      txtCliente.Caption = IIf(IsNull(RsCabecera!Corp), "", RsCabecera!Corp)
      txtReferencia.Caption = IIf(IsNull(RsCabecera!Ref), "", RsCabecera!Ref)
      txtDireccion.Caption = IIf(IsNull(RsCabecera!Dir), "", RsCabecera!Dir)
      txtTelefono.Caption = IIf(IsNull(RsCabecera!Tel), "", RsCabecera!Tel)
      txtObservacion.Caption = IIf(IsNull(RsCabecera!tObservacion), "", RsCabecera!tObservacion)
   End If
      
   RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
   
   If RsDetalle.RecordCount <> 0 Then
      RsDetalle.MoveFirst
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = RsDetalle!tItem
   Else
      sProducto = ""
      sItem = ""
      
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
   End If
  
   'Imprime el Total
   txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   Screen.MousePointer = vbDefault
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsT3 = Nothing
    Set RsCabecera = Nothing
    Set RsDetalle = Nothing
    Set RsImpresion = Nothing
    Set RsGrupo = Nothing
    Set RsSubGrupo = Nothing
    Set RsProducto = Nothing
    Set RsArea = Nothing
    Set RsPropiedad = Nothing
    Set RsTipoDocumento = Nothing
    Set frmVenta = Nothing
End Sub

Private Sub grdCabecera_HeadClick(ByVal ColIndex As Integer)
    Screen.MousePointer = vbHourglass
    RsCabecera.Sort = grdCabecera.Columns(ColIndex).DataField & " ASC"
    RsCabecera.Requery
    Screen.MousePointer = vbDefault
End Sub

Public Sub ActivaCabecera(Activa As Boolean)
   Dim i As Integer
   cmdCabecera(0).Visible = Activa
   cmdCabecera(1).Visible = Activa
   cmdCabecera(2).Visible = Activa
   cmdCabecera(4).Visible = Activa
   
   For i = 0 To 4
       cmdDetalle(i).Visible = Activa
       cmdNavegar(i).Visible = Activa
   Next i
   
   If lPrinter Then
      cmdOpcion(3).Visible = Activa
      cmdOpcion(4).Visible = Activa
   End If
   
   cmdOpcion(7).Visible = Activa
   cmdOpcion(13).Visible = Activa
   cmdNavegar(5).Visible = Activa
   
   For i = 6 To 11
       cmdNavegar(i).Visible = Activa
   Next i
   cmdDetalle(6).Visible = Activa

End Sub

Public Sub ActivaFrame(Frame As Frame)
    'fraCabecera.Visible = False
    fraDetalle.Visible = False
    fraPlato.Visible = False
    fraEliminacion.Visible = False
    fraPropiedad.Visible = False
    Frame.Visible = True
End Sub

Public Sub ActivaOpcion(Activa As Boolean)
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(1).Enabled = Not Activa
    cmdOpcion(2).Enabled = Not Activa
    If lPrinter Then
       cmdOpcion(3).Enabled = Activa
       cmdOpcion(4).Enabled = Activa
    End If
    cmdOpcion(5).Enabled = Activa
    cmdOpcion(6).Enabled = Activa
    cmdOpcion(7).Enabled = Activa
    cmdOpcion(8).Visible = Activa
    cmdOpcion(10).Enabled = Activa
    cmdOpcion(11).Visible = Activa

    cmdOpcion(13).Enabled = Activa
    cmdOpcion(14).Enabled = Activa
    cmdOpcion(16).Enabled = Activa
    
    cmdTipoDocumento(1).Enabled = Activa
    cmdTipoDocumento(2).Enabled = Activa
    cmdTipoDocumento(3).Enabled = Activa
    cmdTipoDocumento(4).Enabled = Activa
    
    cmdNavegar(0).Visible = Activa
    cmdNavegar(1).Visible = Activa
    cmdNavegar(2).Visible = Activa
    cmdNavegar(3).Visible = Activa
    cmdNavegar(4).Visible = Activa
    cmdNavegar(5).Visible = Activa
    cmdNavegar(6).Visible = Activa
    cmdNavegar(7).Visible = Activa
    cmdNavegar(8).Visible = Activa
    cmdNavegar(9).Visible = Activa
    cmdNavegar(10).Visible = Activa
    cmdNavegar(11).Visible = Activa
End Sub

Public Sub InsertaProducto(xCantidad As Double)
    Dim PrecioVenta As Double
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    
    If RsDetalle.RecordCount = 0 Then
       sItem = "001"
    Else
       sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
    End If
    
    'Precios con Recargos / Descargos por Tipo de Pedido
    nRecargo = 0
    nDescuento = 0
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
    
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
        
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nOficial = IIf(IsNull(RsProducto!nPrecioVenta), 0, RsProducto!nPrecioVenta) + (nDelivery * IIf(IsNull(RsProducto!nPrecioVenta), 0, RsProducto!nPrecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nOficial = IIf(IsNull(RsProducto!nPrecioVenta), 0, RsProducto!nPrecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nPrecioVenta), 0, RsProducto!nPrecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    Else
       nOficial = IIf(IsNull(RsProducto!nPrecioVenta), 0, RsProducto!nPrecioVenta)
    End If
    
    nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
    
    'Busca Oferta
    Dim sCriterio As String
    Dim nOferta As Double
    Dim lAcumulable As Boolean
    sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1"
    sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(Date, vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(Date, "yyyy/MM/dd 00:00") & "')) and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'"
    sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(Date, "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(Date, "yyyy/mm/dd") & "')) "
    If sTipoPedido = "01" Then
       sCriterio = sCriterio & " and lLocal=1"
    ElseIf sTipoPedido = "02" Then
       sCriterio = sCriterio & " and lDelivery=1"
    Else
       sCriterio = sCriterio & " and lLlevar=1"
    End If
    
    Isql = "select * from TOFERTA where " & sCriterio
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)

    'Inserta descuento
    nOferta = 0
    lAcumulable = False
    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
       RsOferta.MoveFirst
       If RsOferta!nRatio > 0 Then
          nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
       Else
          nOferta = nOficial - IIf(IsNull(RsOferta!nMonto), 0, RsOferta!nMonto)
       End If
    End If
    
    If xDescuento <> 0 And RsProducto!lDescuento Then
       If RsOferta.RecordCount > 0 Then
          If RsOferta!lAcumulable Then
             nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
             nDescuento = nOficial - nPVenta
          Else
             nPVenta = nOficial - nOferta
             nDescuento = nOficial - nPVenta
          End If
       Else
          nPVenta = nOficial - (nOficial * xDescuento / 100)
          nDescuento = nOficial - nPVenta
       End If
    Else
       nPVenta = nOficial - nOferta
       nDescuento = nOficial - nPVenta
    End If
        
    nValor = 1 + (nValor / 100)
    nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
    nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
    nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    
    Isql = "insert into DPEDIDO " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nPrecioCosto, tEstadoItem, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'" & sTipoPedido & "', '" & sItem & "', '" & sProducto & "', " _
                   & "'" & IIf(IsNull(RsProducto!tGrupo), "", RsProducto!tGrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lcombinacion, -1, 0) & "," _
                   & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                   & IIf(IsNull(RsProducto!nPrecioCosto), 0, RsProducto!nPrecioCosto) & "," _
                   & "'N', '" & sComanda & "', getDate(), '', '" & sUsuario & "', " _
                   & IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo) & ", " _
                   & IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto) & ", " _
                   & IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra) & ", " _
                   & IIf(IsNull(RsProducto!nOrden), 0, RsProducto!nOrden) & ")"
                   
    Cn.Execute Isql
    RsDetalle.Requery
    
    txtSuma.Caption = Format(Val(txtSuma.Caption) + (nPVenta * xCantidad), "###,###,###,##0.00")
    RsDetalle.MoveLast
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaProducto()
   If Not RsDetalle.EOF Then
      fraPrecio.Caption = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      cmdPrecio.Enabled = IIf(IsNull(RsDetalle!lModificable), False, RsDetalle!lModificable)

      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sGrupo = IIf(IsNull(RsDetalle!tCodigoGrupo), "", RsDetalle!tCodigoGrupo)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = IIf(IsNull(RsDetalle!tItem), "001", RsDetalle!tItem)
      sCortesia = IIf(IsNull(RsDetalle!tCortesia), "", RsDetalle!tCortesia)
    
      nOficial = IIf(IsNull(RsDetalle!nPrecioOficial), 0, RsDetalle!nPrecioOficial)
      
      nDescuento = IIf(IsNull(RsDetalle!nDescuento), 0, RsDetalle!nDescuento)
      nRecargo = IIf(IsNull(RsDetalle!nRecargo), 0, RsDetalle!nRecargo)
      nPBase = IIf(IsNull(RsDetalle!nPrecioNeto), 0, RsDetalle!nPrecioNeto)
      nImpuesto1 = IIf(IsNull(RsDetalle!nprecioImpuesto1), 0, RsDetalle!nprecioImpuesto1)
      nImpuesto2 = IIf(IsNull(RsDetalle!nprecioImpuesto2), 0, RsDetalle!nprecioImpuesto2)
      nImpuesto3 = IIf(IsNull(RsDetalle!nprecioImpuesto3), 0, RsDetalle!nprecioImpuesto3)
      nPVenta = IIf(IsNull(RsDetalle!nPrecioVenta), 0, RsDetalle!nPrecioVenta)
      nCantidad = IIf(IsNull(RsDetalle!nCantidad), 0, RsDetalle!nCantidad)
    
      txtOficial.Caption = Format(nOficial, "###,###,###,##0.00")
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,###,##0.00")
      txtCantidad.Caption = Format(nCantidad, "##,##0.00")
      txtVenta.Caption = Format(nPVenta * nCantidad, "###,###,###,##0.00")
     
      If nOficial = 0 Then
         txtDPorcentaje.Caption = "0.00"
         txtRPorcentaje.Caption = "0.00"
      Else
         txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
         txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
      End If
        
      txtCortesia.Caption = IIf(IsNull(RsDetalle!Cortesia), "", RsDetalle!Cortesia)
      txtObserva.Caption = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      lblObservacion.Caption = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
                    
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
      AsignaPropiedad
   End If
End Sub

Public Sub CalculaPrecio()
    Acumulado = 0
    If nPVenta = 0 Then
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nRecargo = 0
       nDescuento = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
       nPVenta = nOficial - nDescuento + nRecargo
       Acumulado = IIf(txtImpuesto1.Caption <> "0.00", Acumulado + nPorcentaje1, Acumulado)
       Acumulado = IIf(txtImpuesto2.Caption <> "0.00", Acumulado + nPorcentaje2, Acumulado)
       Acumulado = IIf(txtImpuesto3.Caption <> "0.00", Acumulado + nPorcentaje3, Acumulado)
       Acumulado = 1 + (Acumulado / 100)
       nImpuesto1 = IIf(txtImpuesto1.Caption <> "0.00", nPVenta / Acumulado * nPorcentaje1 / 100, 0)
       nImpuesto2 = IIf(txtImpuesto2.Caption <> "0.00", nPVenta / Acumulado * nPorcentaje2 / 100, 0)
       nImpuesto3 = IIf(txtImpuesto3.Caption <> "0.00", nPVenta / Acumulado * nPorcentaje3 / 100, 0)
       nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
       txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
       txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
    End If
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    
    txtNeto.Caption = Format(nPBase, "###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
End Sub

Public Sub Impuesto()
   Label1(10).Caption = sImpuesto1 & " :"
   Label1(11).Caption = sImpuesto2 & " :"
   Label1(12).Caption = sImpuesto3 & " :"
   
   Label1(10).Visible = IIf(sImpuesto1 = "", False, True)
   Label1(11).Visible = IIf(sImpuesto2 = "", False, True)
   Label1(12).Visible = IIf(sImpuesto3 = "", False, True)
   
   txtImpuesto1.Visible = IIf(sImpuesto1 = "", False, True)
   txtImpuesto2.Visible = IIf(sImpuesto2 = "", False, True)
   txtImpuesto3.Visible = IIf(sImpuesto3 = "", False, True)
   
   cmdImpuesto(0).Caption = sImpuesto1
   cmdImpuesto(1).Caption = sImpuesto2
   cmdImpuesto(2).Caption = sImpuesto3
   
   cmdImpuesto(0).Visible = IIf(sImpuesto1 = "", False, True)
   cmdImpuesto(1).Visible = IIf(sImpuesto2 = "", False, True)
   cmdImpuesto(2).Visible = IIf(sImpuesto3 = "", False, True)
End Sub

Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   
   Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nDescuento = " & nDescuento & ", " & _
           "nRecargo = " & nRecargo & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & ", " & _
           "tCortesia = '" & sCortesia & "' " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sItem & "'"
           
   txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       RsPropiedad.Filter = "tOperador = '" & RsOperador!Codigo & "' and tProducto='" & sProducto & "'"
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sProducto & "'"
    End If
    AsignaComando 20, RsPropiedad, cmdPropiedad()

    For i = 1 To 20
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Caption = ""
    RsProductoPropiedad.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
    If Not RsProductoPropiedad.EOF Then
       RsProductoPropiedad.MoveFirst
       Do While Not RsProductoPropiedad.EOF
          For i = 1 To 20
              If cmdPropiedad(i).Caption = RsProductoPropiedad!Descripcion And RsOperador!Descripcion = RsProductoPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          lblResumen.Caption = lblResumen.Caption & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ", "
          RsProductoPropiedad.MoveNext
       Loop
    End If
End Sub

Private Sub grdCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsCabecera.RecordCount > 0 And RsCabecera.AbsolutePosition > 0 Then
      AsignaPedido
   End If
   On Error Resume Next
   txtBarra.SetFocus
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsCabecera.RecordCount > 0 Then
      If fraPropiedad.Visible = True Then
         nPos = RsDetalle.AbsolutePosition
         RsDetalle.Requery
         RsDetalle.AbsolutePosition = nPos
      End If
      AsignaProducto
      
      If lPropiedad Then
         lPropiedad = False
         cmdDetalle_Click (4)
      End If
   End If
   On Error Resume Next
   txtBarra.SetFocus
   
End Sub


Public Sub LimpiaCabecera()
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
      txtFecha.Caption = ""
      txtTipoPedido.Caption = ""
      txtCliente.Caption = ""
      txtTelefono.Caption = ""
      txtDescuento.Caption = "0.00"
      txtDireccion.Caption = ""
      txtObservacion.Caption = ""
      fraPlato.Visible = True
      txtSuma.Caption = "0.00"
   End If
End Sub

Private Sub txtBarra_KeyPress(KeyAscii As Integer)
' Solo Sirve Para Chincha
' Productos por Lectora de barras
   
   If KeyAscii = 13 And txtBarra.Text <> "" Then
      If RsCabecera.RecordCount <> 0 Then
      RsProducto.Filter = adFilterNone
         RsProducto.MoveFirst
         RsProducto.Find "tbarra = '" & Trim(txtBarra.Text) & "'"
         If Not RsProducto.EOF Then
            sProducto = RsProducto!Codigo
            sComanda = IIf(sTemp = "", "", Mid("0000000000", 1, 10 - Len(Trim(sTemp))) & Trim(sTemp))
                        
            InsertaProducto 1
            
            If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
               lPropiedad = True
            End If
            
         Else
            MsgBox "Producto no encontrado", vbCritical, sMensaje
         End If
         RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      End If
      txtBarra.Text = ""
      sProducto = ""
   End If

End Sub

Private Sub CalculaDescuento()
If RsDetalle.RecordCount <> 0 Then
   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      Dim sCriterio As String
      Dim lAcumulable As Boolean
      Dim nOferta As Double
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(Date, vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(Date, "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(Date, "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(Date, "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      
      lAcumulable = True
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100)
        
         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
         Acumulado = 1 + (Acumulado / 100)
        
         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3

         Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

End Sub
Public Sub EliminaCabecera()
   Dim i As Integer
   Screen.MousePointer = vbHourglass
                         
   Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'   where tCodigoPedido ='" & sPedido & "'"
   Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
   Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
     
   RsCabecera.Requery
   RsDetalle.Requery
   RsProductoPropiedad.Requery
   AsignaPedido
   ActivaOpcion True
   
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!Codigo
      ActivaFrame fraPlato
   End If
   Screen.MousePointer = vbDefault
End Sub

Public Sub EliminaItem()
    Dim xMax As String
    Dim sMotivo As String
        
    xMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
    xMax = Lib.Correlativo(xMax, 3)
    Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
           "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
           "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion) " & _
           "select '" & sPedido & "' as tCodigoPedido, '" & xMax & "' as tItem, tCodigoProducto, tCodigoGRupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tComanda, lImprime, " & _
           "'" & RsCabecera!tUsuario & "' as tUsuario, cast('" & Format(RsCabecera!fFecha, "yyyy-MM-dd hh:nn:ss") & "' as smalldatetime) as fRegistro, " & _
           "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
           "tObservacion, '" & sDescrip & "' as tObservacion, tEstadoItem, lImprimeArea, tArea, '" & sCodigo & "' " & _
           "from DPEDIDO " & _
           "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    Cn.Execute Isql
        
    Cn.Execute "delete from TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
    Cn.Execute "delete from DPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    
    RsProductoPropiedad.Requery
    RsDetalle.Requery
    If RsDetalle.RecordCount = 0 Then
       txtSuma.Caption = "0.00"
    Else
       txtSuma.Caption = Format(txtSuma.Caption - grdDetalle.Columns(3).Text * nPVenta, "###,###,###,##0.00")
    End If
                
    If Not RsDetalle.EOF Then
       RsDetalle.MoveLast
       sItem = RsDetalle!tItem
    Else
       sProducto = ""
       sItem = ""
              
       ActivaFrame fraPlato
       ActivaCabecera True
       ActivaOpcion True
    End If
End Sub

Public Sub CambiaPrecio(Valor As Double)
    nPVenta = Val(Valor)
    nOficial = nPVenta
    
    Acumulado = 0
    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
    Acumulado = 1 + (Acumulado / 100)
    
    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
    
    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    
    Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sItem & "'"
           
    txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
End Sub

