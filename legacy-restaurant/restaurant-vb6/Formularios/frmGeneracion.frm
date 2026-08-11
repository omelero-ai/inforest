VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form frmGeneracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Documentos"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmGeneracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
      Left            =   11760
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   2700
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Facturación Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   37
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmGeneracion.frx":038A
         Top             =   855
         Width           =   240
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a XXXX."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1170
         TabIndex        =   36
         Top             =   870
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo XXXX almacenado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   35
         Top             =   1155
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmGeneracion.frx":059D
         Top             =   1155
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmGeneracion.frx":07B0
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmGeneracion.frx":0AF2
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a XXXXX......."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   210
         TabIndex        =   38
         Top             =   435
         Width           =   5910
      End
   End
   Begin VB.Frame fraDato 
      Caption         =   " Datos del Documento "
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
      Height          =   1335
      Left            =   60
      TabIndex        =   6
      Top             =   3660
      Width           =   11670
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   11520
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   120
         X2              =   11520
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label txtDescuentoTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1770
         TabIndex        =   33
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Total :"
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   32
         Top             =   915
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Motivo del Descuento :"
         Height          =   195
         Index           =   3
         Left            =   3765
         TabIndex        =   31
         Top             =   915
         Width           =   1650
      End
      Begin VB.Label txtDescuento 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5520
         TabIndex        =   30
         Top             =   870
         Width           =   2625
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Index           =   6
         Left            =   9315
         TabIndex        =   25
         Top             =   915
         Width           =   450
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9840
         TabIndex        =   24
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descuento del Documento :"
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   12
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Neto del Documento :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   375
         Width           =   1560
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto del Documento :"
         Height          =   195
         Index           =   5
         Left            =   8100
         TabIndex        =   10
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label txtNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1770
         TabIndex        =   9
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label txtDImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   9840
         TabIndex        =   7
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " División por Montos "
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
      Height          =   1845
      Left            =   6420
      TabIndex        =   19
      Top             =   5040
      Width           =   5325
      Begin VB.CommandButton cmdOpcion 
         Height          =   540
         Index           =   7
         Left            =   3840
         Picture         =   "frmGeneracion.frx":0E34
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Monto del Documento"
         Height          =   615
         Index           =   6
         Left            =   3840
         TabIndex        =   26
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto total"
         Height          =   195
         Left            =   1110
         TabIndex        =   28
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label txtMontoTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Diferencia"
         Height          =   195
         Left            =   1050
         TabIndex        =   23
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto del Documento Dividido"
         Height          =   435
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label txtDiferencia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   330
         Left            =   2160
         TabIndex        =   21
         Top             =   825
         Width           =   1455
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   330
         Left            =   2160
         TabIndex        =   20
         Top             =   405
         Width           =   1455
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   " Cliente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   60
      TabIndex        =   15
      Top             =   6900
      Width           =   10305
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   17
         Left            =   180
         TabIndex        =   16
         Top             =   375
         Width           =   1275
      End
      Begin VB.Label txtRazonS 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   480
         Width           =   8280
      End
   End
   Begin VB.Frame fraTipoDocumento 
      Caption         =   " Tipo de Documento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   60
      TabIndex        =   13
      Top             =   5040
      Width           =   6285
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   330
         Width           =   1320
      End
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
      Index           =   15
      Left            =   10455
      Picture         =   "frmGeneracion.frx":13BE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7545
      Width           =   1275
   End
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "Emitir"
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
      Left            =   10455
      Picture         =   "frmGeneracion.frx":14B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6930
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   795
      Index           =   13
      Left            =   11055
      Picture         =   "frmGeneracion.frx":15B2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2715
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   795
      Index           =   12
      Left            =   11055
      Picture         =   "frmGeneracion.frx":1E7C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   795
      Index           =   11
      Left            =   11055
      Picture         =   "frmGeneracion.frx":2746
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1125
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   795
      Index           =   10
      Left            =   11055
      Picture         =   "frmGeneracion.frx":3010
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   330
      Width           =   735
   End
   Begin TrueOleDBGrid80.TDBGrid grdDocumento 
      Height          =   3600
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   6350
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
      Caption         =   "Documentos por Generar"
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
   Begin VB.Image ImageIntegracion 
      Height          =   735
      Left            =   9240
      Top             =   8400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imageHash 
      Height          =   375
      Left            =   11760
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11400
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmGeneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Option Explicit
Dim numeroSerieImpresora As String 'SUNAT
Dim codigoImpresora As String 'SUNAT
Dim lAplicablePedido As Boolean
Dim rstFuente As ADODB.Recordset
Dim nCantidad As Integer
Dim wCortesia As Double
Dim sConsumo As String
Dim nMontoPedidoFacturar As Double
'imprime dni
Dim xTipoDocSeleccionado As String
Dim RsDocumento As Recordset
Dim RsTotales As Recordset
Dim RsTipoDocumento As Recordset
Dim RsImpresion As Recordset
Dim RsDetalle As Recordset
Dim RsSuma As Recordset
Dim RsT3 As Recordset
Dim RsArea As Recordset

'validacionMontoMInimo
Dim nMontoPedidoFacturarMInimo As Double
'validacionMontoMInimo


Dim sCabecera As String
Dim sDetalle As String
Dim mTipo() As String
Dim mConsumo() As String
Dim nPos As Integer
Dim nNeto As Double
Dim nPrecioOficial As Double
Dim nTotal As Double
Dim xImpuesto1 As Double
Dim xImpuesto2 As Double
Dim xImpuesto3 As Double
Dim nVeces As Integer
Dim Acumulado As Double
Dim lImp1 As Boolean
Dim lImp2 As Boolean
Dim lImp3 As Boolean
Dim lImprimeAlternativa As Boolean
Dim sCodigoDescuento As String
'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tcodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String
Dim scodFact As String
Dim ntotdoc As Double

'FACTURACION_E_PERU
'Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String
Dim cadenaCodigoHash As String

Dim TimpresionDolaresDelivery  As Boolean


Private Sub cmdEmitir_Click()
    Dim i As Integer
    tAutorizacion = ""
    tcodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    '------VALIDA CORREO----------
    Dim sTipoDocum As String
    Dim lValidaEmail As Boolean
    Dim sEmail As String
    Dim sClienteFact As String
    
    'FACTURACION OFISIS
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    
    Dim rdi As Integer
    Dim RsTotalDocumento As Recordset
    
    lblPaso1.Visible = False
    lblPaso2.Visible = False
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    FrameFeSpring.Visible = False
    
    If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
    Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
    End If
    
    Select Case pais
        Case "001" 'Bolivia
                tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                If tAutorizacion <> "" And tDosificacion <> "" Then
                    Else
                    MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                    Exit Sub
                End If
        Case "002" 'Ecuador
               tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
        Case Else 'Peru
            tAutorizacion = ""
            tcodigoControl = ""
            tDosificacion = ""
    End Select
         
    If Not (RsDocumento.EOF Or RsDocumento.BOF) Then
      RsDocumento.MoveFirst
      For i = 1 To nCantidad
            If RsDocumento.Fields("nventa") = 0 Then
                MsgBox "Los documentos deben tener un monto"
                Exit Sub
            End If
          RsDocumento.MoveNext
       Next
    End If
    
    
    If Not (RsDocumento.EOF Or RsDocumento.BOF) Then
        RsDocumento.MoveFirst
        For i = 1 To nCantidad
        nMontoPedidoFacturar = 0
        nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsDocumento!tTipoDocumento & "'", Cn)
        If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar < RsDocumento!nVenta Then
            MsgBox "Revisar los Montos y los Topes de los Tipos de Documentos", vbInformation, sMensaje
            wEnter = False
            Exit Sub
        End If
        End If
        RsDocumento.MoveNext
        Next
    End If
    
    TimpresionDolaresDelivery = False
            '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    Dim sTexto As String
    Dim nActual As Double
    
    If sModulo = "PVCORP" Then
       grdDocumento.MoveFirst
       'Chequea Consistencia
       For i = 1 To nCantidad
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipoDocumento & "'")
           If Not RsTipoDocumento.EOF Then
          
           Else
             MsgBox "Error: Documento sin definición", vbCritical, sMensaje
             Exit Sub
           End If
           grdDocumento.MoveNext
       Next
            
       RsDocumento.MoveFirst
       For i = 1 To nCantidad
           xCantidad = i
           Facturar RsTipoDocumento, RsDocumento!tTipoDocumento, "Division"
       Next i
       Cn.Execute "delete from " & sTemporal
       Screen.MousePointer = vbDefault
       Unload Me
    Else
        Dim sSerie As String
        Dim sCorrela As String
        Dim sPrefijo As String
        Dim sTipoDocumento As String
        'Dim sDocumento As String
        Dim j As Integer
        grdDocumento.MoveFirst
            
        Dim xNeto As Double
        Dim xImpuesto1 As Double
        Dim xImpuesto2 As Double
        Dim xImpuesto3 As Double
        Dim xVenta As Double
        Dim xTVenta As Double
        Dim xDescuento As Double
        Dim xRecargo As Double
        Dim xOficial As Double
    
        If val(str(txtDiferencia.Caption)) <> 0 Then
           MsgBox "Existen diferencia entre El monto total y la suma de las divisiones", vbExclamation, sMensaje
           Exit Sub
        End If
            
        'Verifica si no ha sido enviado el pedido
        If lPrinter And lObligaPrinter Then
           i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
           If i > 0 Then
              MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
              Exit Sub
           End If
        End If
    
        'Chequea Consistencia
        For i = 1 To nCantidad
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipoDocumento & "'")
           If Not RsTipoDocumento.EOF Then
                If RsTipoDocumento!Cliente And (val(RsTipoDocumento!Monto) <= nTotal Or RsTipoDocumento!Monto = 0) And IsNull(RsDocumento!Cliente) Then
                   MsgBox "Error: Documento sin Razon Social", vbCritical, sMensaje
                   Exit Sub
                End If
              
               sTipoDocum = RsDocumento!tTipoDocumento
            
               lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
               sClienteFact = Calcular("Select MAX(Codigo) As codigo From vCLIENTE where Descripcion ='" & RsDocumento!Cliente & "' ", Cn)
               If lValidaEmail = True And IsNull(RsTipoDocumento!Cliente) Then
                  sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sClienteFact & "' ", Cn)
                  If sEmail = "" Then
                     MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                     Exit Sub
                  End If
               End If
                If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True And IsNull(RsTipoDocumento!Cliente) Then
                    Dim TempUbigeo As String
                    Dim TempUrbaniza As String
                    TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sClienteFact & "' ", Cn)
                    TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sClienteFact & "' ", Cn)
                    If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                        MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                        Exit Sub
                    End If
                End If
           Else
              MsgBox "Error: Documento sin definición", vbCritical, sMensaje
              Exit Sub
           End If
           grdDocumento.MoveNext
        Next
        
        'Areas
        Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
    
        Screen.MousePointer = vbHourglass
        wEnter = True
        sDetalle = dbTemporal(sCaja, 19, "tDocumento", "nVarChar(20)", _
                                         "tItem", "nVarChar(3)", _
                                         "tCodigoPedido", "nVarChar(10)", _
                                         "tCodigoProducto", "nVarChar(7)", _
                                         "nPrecioNeto", "Float", _
                                         "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                         "nPrecioVenta", "Float", _
                                         "nPrecioOficial", "Float", _
                                         "nRecargo", "Float", _
                                         "nDescuento", "Float", _
                                         "nCantidad", "Float", _
                                         "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                         "nVenta", "Float", _
                                         "tGuia", "nVarchar(15)", _
                                         "nDescuentoNeto", "Float")
               
        'Calcula los Porcentajes
        ReDim nPorcentaje(nCantidad) As Double
        grdDocumento.MoveFirst
        If CmdOpcion(6).Enabled Then
            For i = 0 To nCantidad - 1
                nPorcentaje(i) = RsDocumento!nVenta * 100 / val(str(txtMontoTotal.Caption))
                RsDocumento.MoveNext
            Next i
        Else
            For i = 0 To nCantidad - 1
                nPorcentaje(i) = 100 / nCantidad
                RsDocumento.MoveNext
            Next i
        End If
               
        'Chequea Descuento
        Dim nTotalDescuento As Double
        nTotalDescuento = CDbl(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tCodigoPedido='" & sPedido & "'", Cn))
        If nTotalDescuento > 0 Then
           Dim nTope As Double
           Dim nTotalMes As Double
           Dim nConsumo As Double
           Dim aplicaTope As Boolean
           
           
           lAplicablePedido = Calcular("select lAplicablePedido as Codigo FROM vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
           nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
           
           If nTope > 0 Then
              If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
                 If nTotalDescuento > nTope Then
                    If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                       CalculaAplicaTope (nTope)
                       MsgBox "Se realizó el Recálculo del Pedido, " & Chr(13) & "Vuelva a entrar a Pago y División de documento para facturar", vbExclamation, sMensaje
                       wEnter = False
                       Screen.MousePointer = vbDefault
                       Unload Screen.ActiveForm
                       Unload Screen.ActiveForm
                    End If
                    Exit Sub
                 End If
              Else
                 nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                      "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
                 
                 If nTotalDescuento + nTotalMes > nTope Then
                    If nTotalDescuento < nTope Then
                       If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                          CalculaAplicaTope (nTope - nTotalMes)
                          MsgBox "Se realizó el Recálculo del Pedido, " & Chr(13) & "Vuelva a entrar a Pago y División de documento para facturar", vbExclamation, sMensaje
                          wEnter = False
                          Screen.MousePointer = vbDefault
                          Unload Screen.ActiveForm
                          Unload Screen.ActiveForm
                       End If
                       Exit Sub
                    Else
                        MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                        Exit Sub
                    End If
                 End If
              End If
           End If
        End If
        sCodigoDescuento = IIf(lAplicablePedido, "", sCodigoDescuento)
        
        'Genera Correlativo
        grdDocumento.MoveFirst
        For i = 0 To nCantidad - 1
            
            'Verifica el Tipo de Documento
            RsTipoDocumento.MoveFirst
            RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipoDocumento & "'")
                    
            sConsumo = RsDocumento!tConsumo
            If RsTipoDocumento!tFormulario <> "01" And Len(Trim(RsDocumento!tConsumo)) = 0 Then
               RsTipoDocumento.Requery
               RsTipoDocumento.MoveFirst
               RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipoDocumento & "'")
               If RsTipoDocumento.EOF Then
                  MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
                  Exit Sub
               End If
               
               'Genera el Numero de documento
               sSerie = RsTipoDocumento!tSerie
               sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
               sPrefijo = RsTipoDocumento!prefijo
               sTipoDocumento = RsTipoDocumento!TTipoEmision
               sDocumento = sPrefijo & sSerie & sCorrela
               sResumen = RsTipoDocumento!lResumen
               Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
               RsTipoDocumento.Requery
                                    
               Set RsDetalle = Lib.OpenRecordset("select * from " & sTemporal & " where Estado='X'", Cn)
               RsDetalle.MoveFirst
               j = 0
               
               'Actualiza Documento Cabecera
               Isql = "Update " & sCabecera & " set tDocumento= '" & sDocumento & "', " & _
                                              "tEstadoDocumento = '01', " & _
                                              "tTipoDocumento = '" & sTipoDocumento & "', " & _
                                              "tCaja = '" & sCaja & "', " & _
                                              "tTurno = '" & sTurno & "', " & _
                                              "tUsuario = '" & Mid(sUsuario, 1, 15) & "', " & _
                                              "tEmision = '" & IIf(Len(Trim(sConsumo)) = 0, "D", "C") & "', " & _
                                              "tConsumo = '" & IIf(Len(Trim(sConsumo)) = 0, "", sConsumo) & "', " & _
                                              "fRegistro = getdate() " & _
                                              "where Item = " & i + 1
               Cn.Execute Isql
               
               Do While Not RsDetalle.EOF
                  j = j + 1
                  If j Mod (nFItem + 1) = 0 Then
                     RsTipoDocumento.Requery
                     RsTipoDocumento.MoveFirst
                     RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipoDocumento & "'")
                     If RsTipoDocumento.EOF Then
                        MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
                        Exit Sub
                     End If
                     
                     'Calcula el total de la cabecera
                     Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, sum(nPrecioOficial) as nPrecioOficial, isnull(sum(nDescuento*nCantidad),0) as nDescuento " & _
                                  " from " & sDetalle & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
                     Cn.Execute "Update " & sCabecera & " Set nNeto = " & RsSuma!nNeto & ", nPrecioImpuesto1 = " & RsSuma!nImpuesto1 & ", nPrecioImpuesto2 = " & RsSuma!nImpuesto2 & ", nPrecioImpuesto3 = " & RsSuma!nImpuesto3 & ", nVenta = " & RsSuma!nVenta & ", nPrecioOficial = " & RsSuma!nPrecioOficial & ", nDescuento= " & RsSuma!nDescuento & "  Where tDocumento ='" & sDocumento & "'"

                     sSerie = RsTipoDocumento!tSerie
                     sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
                     sPrefijo = RsTipoDocumento!prefijo
                     sTipoDocumento = RsTipoDocumento!TTipoEmision
                     sDocumento = sPrefijo & sSerie & sCorrela
                     sResumen = RsTipoDocumento!lResumen
                     Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                     RsTipoDocumento.Requery
                                      
                     'Inserta Documento Cabecera
                     Isql = "insert into " & sCabecera & " ( tDocumento, tEstadoDocumento, tTipoDocumento, fFecha, tCaja, nNeto, nRecargo, nDescuento, nPrecioImpuesto1, nPrecioImpuesto2, " & _
                                                         "   nPrecioImpuesto3, nVenta, nPrecioOficial, tCodigoCliente, tTurno, tConsumo, tUsuario, tEmision, fRegistro) " & _
                                                    "Values ('" & sDocumento & "', '01', " & _
                                                    "'" & sTipoDocumento & "', getdate(), " & _
                                                    "'" & sCaja & "', " & _
                                                          RsDocumento!nNeto & ", " & _
                                                          RsDocumento!nRecargo & ", " & _
                                                          RsDocumento!nDescuento & ", " & _
                                                          RsDocumento!nprecioImpuesto1 & ", " & _
                                                          RsDocumento!nprecioImpuesto2 & ", " & _
                                                          RsDocumento!nprecioImpuesto3 & ", " & _
                                                          RsDocumento!nVenta & ", " & _
                                                          RsDocumento!nPrecioOficial & ", " & _
                                                    "'" & RsDocumento!tCodigoCliente & "', " & _
                                                    "'" & sTurno & "', " & _
                                                    "'" & RsDocumento!tConsumo & "', " & _
                                                    "'" & Mid(sUsuario, 1, 15) & "', " & _
                                                    "'" & IIf(Len(Trim(sConsumo)) = 0, "D", "C") & "', " & _
                                                    "getdate()) "
                     Cn.Execute Isql
                  End If
                                
                  Isql = "insert into " & sDetalle & _
                         "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                         "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                         "Values ('" & sDocumento & "', " _
                                 & "'" & RsDetalle!tItem & "', " _
                                 & "'" & RsDetalle!tCodigoPedido & "', " _
                                 & "'" & RsDetalle!tCodigoProducto & "', " _
                                 & RsDetalle!nPrecioNeto & ", " _
                                 & RsDetalle!nprecioImpuesto1 & ", " _
                                 & RsDetalle!nprecioImpuesto2 & ", " _
                                 & RsDetalle!nprecioImpuesto3 & ", " _
                                 & RsDetalle!nprecioVenta & ", " _
                                 & RsDetalle!nRecargo & ", " _
                                 & RsDetalle!nDescuento & ", " _
                                 & RsDetalle!nCantidad * nPorcentaje(i) / 100 & ", " _
                                 & RsDetalle!nPrecioOficial & ", " _
                                 & RsDetalle!nImpuesto1 * nPorcentaje(i) / 100 & ", " _
                                 & RsDetalle!nImpuesto2 * nPorcentaje(i) / 100 & ", " _
                                 & RsDetalle!nImpuesto3 * nPorcentaje(i) / 100 & ", " _
                                 & RsDetalle!nVenta * nPorcentaje(i) / 100 & ") "
                  Cn.Execute Isql
                  RsDetalle.MoveNext
               Loop
               
               Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, sum(nPrecioOficial) as nPrecioOficial " & _
                " from " & sDetalle & " where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
               Cn.Execute "Update " & sCabecera & " Set nNeto = " & RsSuma!nNeto & ", nPrecioImpuesto1 = " & RsSuma!nImpuesto1 & ", nPrecioImpuesto2 = " & RsSuma!nImpuesto2 & ", nPrecioImpuesto3 = " & RsSuma!nImpuesto3 & ", nVenta = " & RsSuma!nVenta & ", nPrecioOficial = " & RsSuma!nPrecioOficial & "  Where tDocumento ='" & sDocumento & "'"

          Else
               'Genera y Actualiza los Numero de Documento
               RsTipoDocumento.MoveFirst
               RsTipoDocumento.Find "tTipoEmision ='" & RsDocumento!tTipoDocumento & "'"
               sSerie = RsTipoDocumento!tSerie
               sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
               sPrefijo = RsTipoDocumento!prefijo
               sTipoDocumento = RsTipoDocumento!TTipoEmision
               sDocumento = sPrefijo & sSerie & sCorrela
               sResumen = RsTipoDocumento!lResumen
               Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & RsDocumento!tTipoDocumento & "' and tCaja ='" & sCaja & "'"
               RsTipoDocumento.Requery
            
               'Inserta Detalle de Documento desde el temporal
               Isql = "insert into " & sDetalle & _
                      "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                      "select '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad * " & nPorcentaje(i) & "/100, nPrecioOficial, nImpuesto1 * " & nPorcentaje(i) & "/100, nImpuesto2 * " & nPorcentaje(i) & "/100, nImpuesto3 * " & nPorcentaje(i) & "/100, nVenta *" & nPorcentaje(i) & "/100 From " & sTemporal & " where Estado = 'X'"
               
               Cn.Execute Isql
              
               Isql = "Update " & sCabecera & " set tDocumento= '" & sDocumento & "', " & _
                                              "tEstadoDocumento = '01', " & _
                                              "tTipoDocumento = '" & sTipoDocumento & "', " & _
                                              "tCaja = '" & sCaja & "', " & _
                                              "tTurno = '" & sTurno & "', " & _
                                              "tUsuario = '" & Mid(sUsuario, 1, 15) & "', " & _
                                              "tEmision = '" & IIf(Len(Trim(sConsumo)) = 0, "D", "C") & "', " & _
                                              "tConsumo = '" & IIf(Len(Trim(sConsumo)) = 0, "", sConsumo) & "', " & _
                                              "fRegistro = getdate() " & _
                                              "where Item = " & i + 1
               Cn.Execute Isql
               
            End If
            grdDocumento.MoveNext
        Next
                                        
        'Actualiza la Cabecera con los totales
        Isql = "SELECT [" & sCabecera & "].tDocumento, [" & sCabecera & "].tTipoDocumento, [" & sCabecera & "].tConsumo, [" & sCabecera & "].nRecargo AS Recargo, [" & sCabecera & "].nDescuento AS Descuento, [" & sCabecera & "].nVenta AS Venta, Sum([" & sDetalle & "].[nVenta]-[nImpuesto1]-[nImpuesto2]-[nImpuesto3]) AS Neto, Sum([" & sDetalle & "].nImpuesto1) AS Impuesto1, Sum([" & sDetalle & "].nImpuesto2) AS Impuesto2, Sum([" & sDetalle & "].nImpuesto3) AS Impuesto3, Sum([" & sDetalle & "].nVenta) AS Total " & _
               "FROM " & sCabecera & " LEFT JOIN " & sDetalle & " ON [" & sCabecera & "].tDocumento = [" & sDetalle & "].tDocumento " & _
               "GROUP BY [" & sCabecera & "].tDocumento, [" & sCabecera & "].tTipoDocumento, [" & sCabecera & "].nRecargo, [" & sCabecera & "].nDescuento, [" & sCabecera & "].nVenta, [" & sCabecera & "].tConsumo "
                       
        Set RsSuma = Lib.OpenRecordset(Isql, Cn)
        RsSuma.MoveFirst
        
        Do While Not RsSuma.EOF
  
           'Actualiza Base de Datos Detalle de Pedido
           If sPrefijo = "0" Then
              Set RsT3 = Lib.OpenRecordset("select tCodigoPedido from DPEDIDO where tDocumento = '" & RsSuma!tDocumento & "'", Cn)
              Cn.Execute "Update DPEDIDO set tFacturado = 'C', tDocumento = '" & RsSuma!tDocumento & "' where tCodigoPedido+tItem in (select tcodigoPedido+tItem  from " & sTemporal & " where Estado='X')"
              If nCantidad > 1 Then
                 Do While Not RsT3.EOF
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1 where tCodigoPedido = '" & RsT3!tCodigoPedido & "'"
                    If Not IsNull(RsT3!tMesa) Then
                       Cn.Execute "Update TMESA set tEstadoMesa ='04' where tCodigoMesa='" & RsT3!tMesa & "'"
                    End If
                    RsT3.MoveNext
                 Loop
              End If
           Else
              Isql = "UPDATE DPEDIDO SET DPEDIDO.tFacturado = 'F', DPEDIDO.tDocumento = " & sDetalle & ".tDocumento " & _
                     "FROM  DPEDIDO INNER JOIN " & sDetalle & " ON DPEDIDO.tCodigoPedido = " & sDetalle & ".tCodigoPedido AND DPEDIDO.tItem = " & sDetalle & ".tItem "
              Cn.Execute Isql
           End If
             
           'PARA NO FISCALES
           Isql = "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & RsSuma!tTipoDocumento & "')  from dpedido where dpedido.tdocumento='" & RsSuma!tDocumento & "'"
           Cn.Execute Isql
           
               Isql = "Update " & sCabecera & " set nNeto= " & RsSuma!Neto & ", " & _
                                                    "nPrecioImpuesto1 = " & RsSuma!IMPUESTO1 & ", " & _
                                                    "nPrecioImpuesto2 = " & RsSuma!IMPUESTO2 & ", " & _
                                                    "nPrecioImpuesto3 = " & RsSuma!IMPUESTO3 & ", " & _
                                                    "nVenta = " & RsSuma!TOTAL & ", " & _
                                                    "nPrecioOficial = " & RsSuma!TOTAL & " " & _
                                                    "where tDocumento = '" & RsSuma!tDocumento & "'"
                Cn.Execute Isql
                
           RsSuma.MoveNext
        Loop
        
        'Cancela Todos los documentos por Cortesia
        Cn.Execute "Update " & sCabecera & " set tEstadoDocumento = '02' where tCortesia <> '' "
            
        'Inserta Detalle del Documento desde el Temporal
        Isql = " insert into DDOCUMENTO(tDocumento, tItem,  tCodigoPedido,  tCodigoProducto,    nPrecioNeto ,nPrecioImpuesto1,  nPrecioImpuesto2,   nPrecioImpuesto3,   nPrecioVenta,   nPrecioOficial, nRecargo,   nDescuento, nCantidad,  nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tGuia,  nDescuentoNeto, nLey1)" & _
               " select " & sDetalle & ".*, case when vproducto.lley1=1 then " & sDetalle & ".ncantidad*" & CDbl(sValorLey1) & " else 0 end  From " & sDetalle & " inner join vproducto on " & sDetalle & ".tcodigoproducto=vproducto.codigo "
        Cn.Execute Isql
    
        Select Case pais
                Case "001" 'Bolivia
                        Dim rstX As New ADODB.Recordset
                        Set rstX = Lib.OpenRecordset("select tdocumento from " & sCabecera & " order by tdocumento", Cn)
                        If Not (rstX.EOF Or rstX.BOF) Then
                                    rstX.MoveFirst
                                    For i = 1 To rstX.RecordCount
                                            scodFact = Calcular("select tcodigocliente as codigo from " & sCabecera & " where tdocumento='" & rstX.Fields(0) & "' ", Cn)
                                            ntotdoc = Calcular("select nventa as codigo from " & sCabecera & " where tdocumento='" & rstX.Fields(0) & "' ", Cn)
                                            tcodigoControl = devuelveCodigoControl(sCaja, Mid(rstX.Fields(0), 7, 9), tAutorizacion, tDosificacion, scodFact, ntotdoc)
                                            Cn.Execute "update " & sCabecera & " set tautorizacion=' " & tAutorizacion & "', tcodigocontrol='" & tcodigoControl & "' where tdocumento='" & rstX.Fields(0) & "'"
                                            
                                            rstX.MoveNext
                                    Next i
                        End If
                Case Is = "002" 'Ecuador
                        Cn.Execute "update " & sCabecera & " set tautorizacion= '" & tAutorizacion & "' , tcodigocontrol='' "
                        
                Case Else 'Peru
                        Cn.Execute "update " & sCabecera & " set tautorizacion='', tcodigocontrol='' "
        
        End Select
        
        'Inserta Cabecera de Documento desde el Temporal
        Isql = "insert into MDOCUMENTO " & _
               "     ( tDocumento, tCodigoCliente, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nVenta, nPropina, nTotal, tEstadoDocumento, tCaja, tCortesia, tTipoDocumento, tTurno, tSalon, tUsuario, tUsuarioAutoriza, fRegistro, tEmision, tConsumo,tautorizacion,tcodigocontrol,FDIACONTABLE,tdescuento, lImpresionMonedaExtranjera, nLey1) " & _
               "select tDocumento, tCodigoCliente, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
               "nVenta, nPropina, nTotal, tEstadoDocumento, tCaja, tCortesia, tTipoDocumento, tTurno, '" & sSalon & "' as tSalon, tUsuario, tUsuario, fRegistro, tEmision, tConsumo,tautorizacion,tcodigocontrol,'" & Format(obtieneDiaContable, "yyyyMMdd") & "' , tdescuento , " & IIf(TimpresionDolaresDelivery, 1, 0) & _
               ", (select round(sum(isnull(nley1,0)),2) from ddocumento where tdocumento= " & sCabecera & ".tdocumento)  From " & sCabecera
        Cn.Execute Isql
        
        Cn.Execute " update mdocumento set nventa=nventa+nley1 where tdocumento in (select tdocumento from " & sCabecera & ") "
        
        wEnter = True
        If lPagoAntesImpresion Then
            Dim RsCantDocumentosTemp As Recordset
            
            'GCAA 16092022
            EstoyDivisionDoc = "1"
            
            'JESUS 06/03/2018
            Set RsCantDocumentosTemp = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento not in('04','02') AND M.TDOCUMENTO NOT IN (SELECT ISNULL(TDOCUMENTO,'') FROM VNOTACREDITO WHERE TESTADODOCUMENTO IN ('02','05'))", Cn)
            Dim rdiTemp As Integer
            For rdiTemp = 0 To RsCantDocumentosTemp.RecordCount - 1
                    sDocumento = RsCantDocumentosTemp!tDocumento
                    
                    Set RsTotalDocumento = Lib.OpenRecordset("select nventa " & _
                                          " from MDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                    nCargo = Round(RsTotalDocumento!nVenta, 2)
                    
                    sFormulario = ""
                    sTipo = ""
                    frmPago.Show vbModal
                    
                    If wEnter = False Then
                        Dim RsCantDocumentos As Recordset
                        ' CAMBIO ELDC
                        Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento, M.tTipoDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento not in('04') AND M.TDOCUMENTO NOT IN (SELECT ISNULL(TDOCUMENTO,'') FROM VNOTACREDITO WHERE TESTADODOCUMENTO IN ('02','05'))", Cn)
                        'Set RsCantDocumentos = Lib.OpenRecordset("select distinct DP.tDocumento, M.tTipoDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento inner join DPEDIDO DP ON D.tCodigoPedido=DP.tCodigoPedido where D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento <> '04' and DP.tFacturado= 'F'", Cn)
                        For rdi = 0 To RsCantDocumentos.RecordCount - 1
                            Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                            Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                            Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & sPedido & "'"
 
                            Dim xUltimoCorrelativo As String
                            xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & RsCantDocumentos!tTipoDocumento & "' and left(tdocumento,1)= left('" & RsCantDocumentos!tDocumento & "',1)", Cn)
                            xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                        
                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & RsCantDocumentos!tTipoDocumento & "' and tCaja ='" & sCaja & "'"
                            RsCantDocumentos.MoveNext
                        Next rdi

                       'GCAA 16092022
                         Cn.Execute "update mpedido set testadopedido='01' where tcodigopedido='" & sPedido & "'"
                         Unload Me
                        Exit Sub
    
                    End If

                    RsCantDocumentosTemp.MoveNext
            Next rdiTemp
            Cn.Execute "update DPEDIDO set tFacturado = 'P' where tCodigoPedido = '" & sPedido & "' and tFacturado = 'F'"
        End If
        
        '-----------------------
        If pais = "000" And lFEpape Then
            Dim RsCantDocumentosTemp0 As Recordset
            Set RsCantDocumentosTemp0 = Lib.OpenRecordset("select distinct m.tDocumento,m.tTipoDocumento,isnull((SELECT lFacturacionElectronica FROM vTipodocumentoImpresora where tTipoEmision=m.tTipoDocumento and tCaja=m.tCaja),0) as lFacturacionElectronica  from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento inner join DPEDIDO DP ON D.tCodigoPedido=DP.tCodigoPedido  where  D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento not in( '04')", Cn)
            Dim rdiTemp0 As Integer
            For rdiTemp0 = 0 To RsCantDocumentosTemp0.RecordCount - 1
             RsTipoDocumento.MoveFirst
             RsTipoDocumento.Find ("tTipoEmision='" & RsCantDocumentosTemp0!tTipoDocumento & "'")
                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                    If Not FacturarTCPIP(2, RsCantDocumentosTemp0!tDocumento, 0) Then
                        MsgBox "NO se pudo enviar el Documento " + RsCantDocumentosTemp0!tDocumento, vbInformation
'                        Dim RsCantDocumentos0 As Recordset
'                        Set RsCantDocumentos0 = Lib.OpenRecordset("select distinct DP.tDocumento, M.tTipoDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento inner join DPEDIDO DP ON D.tCodigoPedido=DP.tCodigoPedido where D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento <> '04' ", Cn)
'                        For rdi = 0 To RsCantDocumentos0.RecordCount - 1
'                            Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos0!tDocumento & "'"
'                            Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos0!tDocumento & "'"
'                            Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos0!tDocumento & "'"
'                            Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & sPedido & "'"
'
'                            Dim xUltimoCorrelativo0 As String
'                            xUltimoCorrelativo0 = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & RsCantDocumentos0!tTipoDocumento & "'", Cn)
'                            xUltimoCorrelativo0 = Right(xUltimoCorrelativo, 9)
'
'                            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo0 & "' where tTipoEmision ='" & RsCantDocumentos0!tTipoDocumento & "' and tCaja ='" & sCaja & "'"
'                            RsCantDocumentos0.MoveNext
'                        Next rdi
'                        Exit Sub
                    End If
                    'If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                         If Not FacturarTCPIP(3, RsCantDocumentosTemp0!tDocumento, 0) Then
                            MsgBox ("La confirmacion ha fallado Favor de Reenviar el Documento, en Correlativo de Documentos..."), vbInformation, sMensaje
                         End If
                    'End If
                End If
                RsCantDocumentosTemp0.MoveNext
            Next
         End If

        'Crea Comandas en Infhotel
        If lInfhotel Then
            Dim xSuma As Double
            Dim PrecioTotal As Double
            Dim sComandaInfhotel As String
            Dim EstadoC As String
            Dim RsTCabecera As New ADODB.Recordset
            Dim rsPuntoVenta As New ADODB.Recordset
            Dim rsTDocumento As New ADODB.Recordset
            
            Isql = "SELECT * from vPedidoCabecera where Codigo = '" & sPedido & "'"
            Set RsTCabecera = Lib.OpenRecordset(Isql, Cn)
            
            Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda From tPuntoVenta where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
            Set rsPuntoVenta = Lib.OpenRecordset(Isql, CnInfhotel)
            If rsPuntoVenta.RecordCount <= 0 Then
                MsgBox "No se a configurado el punto de venta en Infhotel", vbExclamation, sMensaje
            Else
                rsPuntoVenta.MoveFirst
            
                RsTCabecera.MoveFirst
    
                RsSuma.Requery
                RsSuma.MoveFirst
                
                RsTipoDocumento.MoveFirst
                RsTipoDocumento.Find ("tTipoEmision='" & RsSuma!tTipoDocumento & "'")
                
                Isql = "select tDocumento, tCortesia from  " & sCabecera & " where tDocumento='" & RsSuma!tDocumento & "'"
                Set rsTDocumento = Lib.OpenRecordset(Isql, Cn)
        
                xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn)
                If RsTipoDocumento!TTipoEmision = "00" Then
                    PrecioTotal = 0
                    EstadoC = "01"
                Else
                    PrecioTotal = xSuma 'RsSuma!Venta
                    EstadoC = "02"
                End If
    
                If Len(Trim(RsTCabecera!tComanda)) = 0 Or IsNull(RsTCabecera!tComanda) Then
                    sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & RsTCabecera!tPuntoVenta & "'", CnInfhotel)
                    sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
                    CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & RsTCabecera!tPuntoVenta & "'"
                    sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!Descripcion, 1, 3))
                    rsPuntoVenta.Requery
                    rsPuntoVenta.MoveFirst
                    rsPuntoVenta.Find "Codigo='" & RsTCabecera!tPuntoVenta & "'"
    
                      'Genero las comandas en Infhotel
                      'Cabecera
                      'CESAR ECUADOR sDocumento 15 Infhotel
                    Isql = "Insert into MCOMANDA " & _
                        "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                        "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                        "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE,tnotapedido) " & _
                        "values('" & sComandaInfhotel & "', '" & RsTCabecera!tPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & xSuma & ", '" & EstadoC & "', " & _
                        "1, '" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', '', '', '" & IIf(RsTipoDocumento!TTipoEmision = "00", Mid(rsTDocumento!tCortesia, 3, 2), "") & "', " & _
                        "'" & sCajaInfhotel & "', '" & IIf(pais = "002", Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 3), RsSuma!tDocumento) & "', '" & xUsuario & "', " & nTC & ", '', '" & RsTCabecera!tPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "', '" & sPedido & "')"
                             '"'" & sCajaInfhotel & "', '" & sDocumento & "', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "', '" & sPedido & "')"
                    CnInfhotel.Execute Isql
                Else
                    sComandaInfhotel = RsTCabecera!tComanda
                    CnInfhotel.Execute "update MCOMANDA set TASIGNACION='" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', TCODIGORESERVA='', TNUMEROHABITACION='', TCLIENTE='', nTotal= " & xSuma & ", tEstado='" & EstadoC & "', tDocumento='" & IIf(pais = "002", Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 3), RsSuma!tDocumento) & "' " & _
                                         "where tComanda ='" & RsTCabecera!tComanda & "' and tPuntoVenta='" & RsTCabecera!tPuntoVenta & "'"
                End If
    
                   'Detalle
                Dim nMovimiento As Integer
                CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & RsTCabecera!tPuntoVenta & "'"
                nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                If RsTipoDocumento!TTipoEmision = "00" Then
                    Isql = "Insert into DCOMANDA " & _
                                      "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                                      "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                                      "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                                      "select '" & sComandaInfhotel & "' as tComanda, '" & RsTCabecera!tPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ",  '" & sPedido & "' as  tNotaPedido, tInfhotel as tCodigoItem, " & _
                                      "0, nCantidad, 0, 0, '" & RsTCabecera!tReserva & "' as tCodigoReserva, '" & RsTCabecera!tHabitacion & "' as tNumeroHabitacion, " & _
                                      "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & IIf(pais = "002", Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 3), RsSuma!tDocumento) & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                                      "where tCodigoPedido ='" & sPedido & "'"
                               CnInfhotel.Execute Isql
                Else
                    Isql = "Insert into DCOMANDA " & _
                                      "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                                      "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                                      "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                                      "select '" & sComandaInfhotel & "' as tComanda, '" & RsTCabecera!tPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ",  '" & sPedido & "' as  tNotaPedido, tInfhotel as tCodigoItem, " & _
                                      "T1.nPrecioVenta as nPrecioUnitario, nCantidad, nVenta as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                                      "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & IIf(pais = "002", Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 3), RsSuma!tDocumento) & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                                      "where  tCodigoPedido ='" & sPedido & "'"
                               CnInfhotel.Execute Isql
                End If
                Cn.Execute "update MPEDIDO set tComanda = '" & sComandaInfhotel & "'  where tCodigoPedido='" & sPedido & "'"
            End If
        End If
        
        'impresion imagen
        Set rstFuente = New ADODB.Recordset
        imageCab.Picture = Nothing
        imagepIE.Picture = Nothing
        Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
        imageCab.DataField = "foto"
        Set imageCab.DataSource = rstFuente
        imagepIE.DataField = "fotoPie"
        Set imagepIE.DataSource = rstFuente
                
        'Imprime Documentos
        RsSuma.Requery
        RsSuma.MoveFirst
        Do While Not RsSuma.EOF
           lImprimeAlternativa = False
           If Len(Trim(RsSuma!tConsumo)) <= 0 And lDescripcionAlternativa = True Then
              If validaImpresionAlternativa(RsSuma!tDocumento) = False Then
                 If MsgBox("Desea imprimir descripción Alternativa " & RsSuma!tDocumento & "? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                    lImprimeAlternativa = True
                 Else
                    lImprimeAlternativa = False
                 End If
              End If
           End If
           

           '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
           ' ELDCQ 15/11/2017
           If lImprimeAlternativa = False Then
                    'Imprime Documentos
                    If lDocumentoAgrupado Then
                        Isql = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',1"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',2"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',3"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',4"
                    End If
            Else
                    If lDocumentoAgrupado Then
                        Isql = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',5"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',6"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',7"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsSuma!tDocumento & "',8"
                    End If
            End If
        
        Dim xxUltimoCorrelativo As String
        lblPaso1.Visible = False
        lblPaso2.Visible = False
        imgProceso(0).Visible = False
        imgProceso(1).Visible = False
        imgProceso(2).Visible = False
        imgProceso(3).Visible = False
        FrameFeSpring.Visible = False

        Dim imprimeDolaDocumentos As String
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & RsSuma!tTipoDocumento & "' and Transporte=0", Cn)
        
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
        'FACTURACION_E_PERU
        'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & RsSuma!tDocumento & "')", Cn)
        xImpresioDE = Mid(RsSuma!tDocumento, 1, 1)
        '---------------------------------------
        
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        'RsImpresion.Requery
    
       If RsImpresion.RecordCount <> 0 Then
            'Configura la Impresora
            RsTipoDocumento.MoveFirst
            RsTipoDocumento.Find ("tTipoEmision='" & RsSuma!tTipoDocumento & "'")
            
            If RsTipoDocumento.EOF Then
               'SUNAT
               numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, "001")
               codigoImpresora = "001"
              
               Imprimir ("001")
               Printer.FontName = sFont
               Printer.FontBold = False
            Else
              'SUNAT
               numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, RsTipoDocumento!timpresora)
               codigoImpresora = RsTipoDocumento!timpresora
               
               Imprimir (RsTipoDocumento!timpresora)
               Printer.FontName = sFont
               Printer.FontBold = False
            End If
            
            'SUNAT
            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & RsSuma!tDocumento & "' "
            
            'JESUS 06/03/2018
            lPagocortesiaAI = IIf(RsImpresion!tCortesia = "", False, True)
          
            'FACTURACION_E_PERU
            If pais = "000" Then
                   If lFacturacionE Then
                   
                        If lFEOfisis Then 'OFISIS
                                '----CABECERA
                                Set oComandoCabeceraOfisis = New clsComando
                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsSuma!tDocumento
    
                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                If Not oComandoCabeceraOfisis.ExecSP Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If

                                '----FIRMA DOCUMENTO OFISIS
                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsSuma!tDocumento
        
                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    
                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                    fDocumento = Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 4, 3) + "-" + CStr(CLng(Mid(RsSuma!tDocumento, 8, 8)))
                                    Sleep 3000
                                    If lImpresionCodigoBarras Then
                                        imageHash.DataField = "foto"
                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                        Set imageHash.DataSource = RsCodigoHash
                                        
                                    ElseIf lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                    Else
                                        Dim RscadenaCodigoHash As Recordset

                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                        End If
                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                    End If
                                End If
                        
                        
                        ElseIf lFESpring Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                 frmMensajeFeSpring.sDocumento = RsSuma!tDocumento
                                 frmMensajeFeSpring.oVenta = 2 ' 2: "Formulario Generacion"
                                 frmMensajeFeSpring.Show vbModal
                                 If frmMensajeFeSpring.lEnvio = False Then
                                    RsSuma.MoveNext
                                 End If
                                
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 4, 3) + Mid(RsSuma!tDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                        If frmMensajeFeSpring.lQrInf Then
                                            Set imageHash.Picture = LoadPicture(ImagenQR(RsSuma!tDocumento))
                                        Else
                                            Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, RsSuma!tDocumento))
                                        End If
                                     Else
                                         If lImpresionCodigoBarras Then
                                             
                                         Else
                                            cadenaCodigoHash = ImagenFeSpring(lQRFE, RsSuma!tDocumento)
                                         End If
                                     End If
                                 End If
                             End If
                        
                        
                        ElseIf lFECarbajal Then
                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                Dim sImporteLetra As String
                                sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                FrameFeSpring.Visible = True
                                lblPaso1.Visible = True
                                lblPaso2.Visible = True
                                imgProceso(0).Visible = False
                                imgProceso(1).Visible = False
                                imgProceso(2).Visible = False
                                imgProceso(3).Visible = False
                                Sleep 1000
                                If Not INSERTAFE_CARVAJAL(RsSuma!tDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                        RevierteDocumento RsSuma!tDocumento, sPedido, sMesa, sCaja
                                        imgProceso(2).Visible = True
                                        imgProceso(3).Visible = True
                                        Sleep 1000
                                        FrameFeSpring.Visible = False
                                        Exit Sub
                                 End If
                                 imgProceso(0).Visible = True
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 4, 3) + Mid(RsSuma!tDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, RsSuma!tDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, RsSuma!tDocumento, 0))
                                         Else
                                             cadenaCodigoHash = ImagenFeCarvajal(2, RsSuma!tDocumento, 0)
                                         End If
                                     End If
                                 End If
                                 imgProceso(1).Visible = True
                                 Sleep 1000
                                 FrameFeSpring.Visible = False
                            End If
                        
                        
                        ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                             cadenaCodigoHash = PapeMatricial
                                         End If
                                     End If
                                 End If
                            End If
                        ElseIf lFEBiz Then
                           If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTA_FE_INFOREST(RsSuma!tDocumento, 1, DateTime.Date) Then '----CABECERA
                                        RevierteDocumento RsSuma!tDocumento, sPedido, sMesa, sCaja
                                        MsgBox "No se pudo Procesar los documentos, favor de verificar los Datos!!", vbInformation, sMensaje
                                        Exit Sub
                                 End If
                                 
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, RsSuma!tDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, RsSuma!tDocumento, 0))
                                         Else
                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, RsSuma!tDocumento, 0)
                                         End If
                                     End If
                                 End If
                            End If
                        ElseIf lFEGesa Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTAFE(RsSuma!tDocumento, "", 1, "") Then '----CABECERA
                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                End If
                                If lQRFE Then
                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, RsSuma!tDocumento, 0))
                                Else
                                    If lImpresionCodigoBarras Then
                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, RsSuma!tDocumento, 0))
                                    Else
                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, RsSuma!tDocumento, 0)
                                    End If
                                End If
                            End If
                        Else 'INFOFACT
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                 If Not INSERTAFE(RsSuma!tDocumento, "", 1, "") Then '----CABECERA
                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                     'Exit Sub
                                 End If
'                                If RsImpresion!Ruc <> "" Then
'                                    If Not INSERTAFE(RsSuma!tDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                        Exit Sub
'                                    End If
'                                End If
                                'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                fDocumento = Mid(RsSuma!tDocumento, 1, 1) + Mid(RsSuma!tDocumento, 4, 3) + Mid(RsSuma!tDocumento, 8, 8)
                                If tCodigoFE = "000" Then
                                    If lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR(RsSuma!tDocumento))
                                    Else
                                        If lImpresionCodigoBarras Then
                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, RsSuma!tDocumento))
                                        Else
                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, RsSuma!tDocumento)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                   End If
                   
            End If

            If lIntVisaNetQR And sTipoDocumento <> "00" Then
                Set ImageIntegracion = Nothing
                If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 1, 1) = True) Then
                    'Sleep (500)
                    Dim Ob As New Scripting.FileSystemObject
                    If Ob.FileExists(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp") Then
                        Set ImageIntegracion.Picture = LoadPicture(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp")
                        Kill lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp"
                    Else
                     Set ImageIntegracion = Nothing
                    End If
                Else
                    MsgBox "No se pudo generar el QR Visa Net, El sistema continuara con el proceso!!!", vbInformation
                    Set ImageIntegracion = Nothing
                End If
            Else
                Set ImageIntegracion = Nothing
            End If

            '---------------------------------------
            If pais = "003" And (Left(RsSuma!tDocumento, 1) = "A" Or Left(RsSuma!tDocumento, 1) = "B") Then 'And Calcular("select isnull(tpuertoserial,'') as codigo from timpresora where tcaja='" & sCaja & "' and timpresora='" & RsTipoDocumento!timpresora & "'", Cn) <> "") Or (pais = "003" And lFETusFacturasAPP = True) Then ' Pais Argentina
                Screen.MousePointer = vbHourglass
                    If lFETusFacturasAPP = True Then
                        If Not (INSERTA_FE_INFOREST_ARGE(RsSuma!tDocumento, 1, DateTime.Date)) Then
                            RevierteDocumento RsSuma!tDocumento, sPedido, sMesa, sCaja
                            MsgBox "No se pudo Procesar los documentos  , favor de verificar los Datos!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        Else
                            'ELDCQ 27/04/2022
                            Dim Termico As String
                             Termico = Calcular("select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento='" & RsSuma!tDocumento & "'", Cn)
                             If Termico <> "" Then
                                Set imageHash.Picture = LoadPicture(CrearImagenQR(Termico))
                                ImprimeFacturaElectronicaARGE RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageIntegracion
                            Else
                                'MsgBox "Por el momento los servicios de AFIP estan caidos o hubo un percance en la facturacion el documento intentara enviarse en uno minutos a mas tardar 1 dia, " & vbNewLine & " favor indicar al cliente que revise su correo electronico para la recepcion de la factura: " & sDocumento, vbInformation
                                Dim MENSAJERROR As String
                                MENSAJERROR = Calcular("select isnull(error,'') as codigo from DOCUMENTO_ARGENTINA where documento='" & RsSuma!tDocumento & "'", Cn)
                                MsgBox MENSAJERROR, vbInformation
                            End If
                        End If
                    End If
                    If lFETusFacturasAPP = False Then
                    ' No facturacion
                        If ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion) = 0 Then
                            RevierteDocumento RsSuma!tDocumento, sPedido, sMesa, sCaja
                            MsgBox "No se pudo Procesar los documentos, favor de verificar los Datos!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                
                Screen.MousePointer = vbDefault
            Else
                If Len(Trim(RsSuma!tConsumo)) > 0 Then
                
                        If RsTipoDocumento!tFormulario = "01" Then
                           If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                           
                                 'FACTURACION ELECTRONICA
                                 If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                     ImprimeFacturaConsumoElectronico RsImpresion, RsSuma!tConsumo, imageHash, RsSuma!tTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                 Else
                                     ImprimeFacturaConsumoT RsImpresion, RsSuma!tConsumo, RsSuma!tTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                 End If
                                  
                                  NFactura = sCorrela
                           Else
                                 'FACTURACION ELECTRONICA
                                 If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                     ImprimeBoletaConsumoElectronico RsImpresion, RsSuma!tConsumo, imageHash, RsSuma!tTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                                 Else
                                     ImprimeBoletaConsumoT RsImpresion, RsSuma!tConsumo, RsSuma!tTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                                 End If
                           End If
                           
                        Else  ' VARIABLE POR CONSUMO
                        
                           If lFacturacionE Then
                           
                                        'FACTURACION_E_PERU
                                        'FORMATO A4
                                        If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                           ImprimeFormatoAConsumo RsSuma!tDocumento
                                           Kill App.path & "\fact.bmp"
                                        Else
                                           ImprimeFormatoAConsumo RsSuma!tDocumento
                                        End If
                           Else
                        
                                        If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                           ImprimeFacturaConsumoN RsImpresion, RsSuma!tConsumo, sEmpresa
                                           NFactura = sCorrela
                                        Else
                                           ImprimeBoletaConsumoN RsImpresion, RsSuma!tConsumo, sEmpresa
                                        End If
                           End If
                           
                           
                        End If
                        
                Else  'FORMATO TICKET
                
                   If RsTipoDocumento!tFormulario = "01" Then
                   
                            If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                            
                                'FACTURACION ELECTRONICA
                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                   ImprimeFacturaElectronica RsImpresion, imageHash, RsSuma!tTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                                Else
                                   ImprimeFacturaT RsImpresion, RsSuma!tTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery, ImageIntegracion
                                End If
                               
                                NFactura = sCorrela
                               
                            ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                
                                If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                   ImprimeCortesia RsImpresion, RsSuma!tTipoDocumento, imageCab, imagepIE
                                End If
                               
                            Else
                            
                                'FACTURACION ELECTRONICA
                                If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                   ImprimeBoletaElectronica RsImpresion, imageHash, RsSuma!tTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                                Else
                                   ImprimeBoletaT RsImpresion, RsSuma!tTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery, ImageIntegracion
                                End If
                               
                            End If
                      
                    Else  'FORMATO VARIABLE
                    
                          If lFacturacionE Then
                                        'FACTURACION_E_PERU
                                        'FORMATO A4
                                        If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                           ImprimeFormatoA RsSuma!tDocumento
                                           Kill App.path & "\fact.bmp"
                                        Else
                                           ImprimeFormatoA RsSuma!tDocumento
                                        End If
                                        
                          Else
                          
                                        If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                           ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                           NFactura = sCorrela
                                        ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                           If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                              ImprimeCortesia RsImpresion, RsSuma!tTipoDocumento, imageCab, imagepIE
                                           End If
                                        Else
                                           ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                        End If
                                        
                          End If
                        
                      
                   End If
                   
                   
                End If
            End If ' Argentina
            
               'CESAR FACTURACION ELECTRONICA
               If pais = "002" Then
                   If lFacturacionE And lFEEcuador = False Then
                       If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                            'CABECERA
                            Dim oComandoCabecera2 As clsComando
                            Set oComandoCabecera2 = New clsComando
                            If Not oComandoCabecera2.CreateCmdSp("USP_InsertaFactDocumento", Cn) Then
                                 Set oComandoCabecera2 = Nothing
                                 Exit Sub
                            End If
                            oComandoCabecera2.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsSuma!tDocumento
            
                            If Not oComandoCabecera2.GetParamOK Then
                                 Set oComandoCabecera2 = Nothing
                                 Exit Sub
                            End If
                            If Not oComandoCabecera2.ExecSP Then
                                 Set oComandoCabecera2 = Nothing
                                 Exit Sub
                            End If
                            'DETALLE
                            Dim oComandoDetalle2 As clsComando
                            Set oComandoDetalle2 = New clsComando
                            If Not oComandoDetalle2.CreateCmdSp("USP_InsertaFactDocumentoDetalle", Cn) Then
                                 Set oComandoDetalle2 = Nothing
                                 Exit Sub
                            End If
                            oComandoDetalle2.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsSuma!tDocumento
        
                            If Not oComandoDetalle2.GetParamOK Then
                                 Set oComandoDetalle2 = Nothing
                                 Exit Sub
                            End If
                            If Not oComandoDetalle2.ExecSP Then
                                 Set oComandoDetalle2 = Nothing
                                 Exit Sub
                            End If
                        End If
                   End If
               End If
               '---------------------------------
            If pais = "002" And lFEEcuador = False Then
                Dim sXML As String
               sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
               GeneraFacturaElectronica sXML, RsSuma!tDocumento
            End If
            
            If lFEEcuador Then
             If INSERTA_FE_INFOREST(RsSuma!tDocumento, 1, DateTime.Now) = False Then
                 MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
             End If
       End If

       End If
       RsSuma.MoveNext
       Set RsImpresion = Nothing
       Loop
       
       If (Calcular("select count(*) as Codigo from dpedido where tcodigopedido='" & sPedido & "' and isnull(tdocumento,'')=''", Cn)) = 0 Then
        Cn.Execute "update mpedido set testadopedido='02' where tcodigopedido='" & sPedido & "'"
       End If
       
       'Borra los platos por facturar
       Cn.Execute "delete from " & sTemporal & " where Estado ='X'"

       Cn.Execute "drop table " & sDetalle
       Screen.MousePointer = vbDefault
       Unload Me
    End If
End Sub


Private Sub ImprimeFormatoAConsumo(ByVal nDocumento As String)

                        Dim ReporteC As New dsrBoletaC
                        
                        If RsTipoDocumento!lImprimeImageCab Then
                           iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                        End If
                        
                        ReporteC.DiscardSavedData
                        ReporteC.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                        
                        If xImpresioDE = "B" Then
                           ReporteC.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                        ElseIf xImpresioDE = "F" Then
                           ReporteC.Text13.SetText "FACTURA ELECTRONICA"
                        End If
                                                            
                        ReporteC.Text8.SetText sRazonSocial
                        ReporteC.ReportTitle = sDireccion
                        ReporteC.Text15.SetText sTelefono
                        ReporteC.Text33.SetText sFax
                        ReporteC.Text16.SetText sRUC
                        ReporteC.Text50.SetText sWeb
                        
                        ReporteC.Text31.SetText RsSuma!tConsumo
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                           ReporteC.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                        ReporteC.Text4.SetText xMontoTexto
                        ReporteC.Text32.SetText xImpresionFE

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = ReporteC
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                        
                        ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ReporteC.PaperOrientation = crPortrait
                        ReporteC.PrintOut False, 1, False, 1, 1
                        '----------------
                                                          
                    
                        If iImagenCab Then
                           Kill App.path & "\cliente.jpg"
                        End If
End Sub

Private Sub ImprimeFormatoA(ByVal nDocumento As String)
                    Dim Reporte As New dsrBoleta

                    If RsTipoDocumento!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                
                    Reporte.DiscardSavedData
                    Reporte.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                    
                    If xImpresioDE = "B" Then
                       Reporte.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       Reporte.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                                                        
                    Reporte.Text8.SetText sRazonSocial
                    Reporte.ReportTitle = sDireccion
                    Reporte.Text15.SetText sTelefono
                    Reporte.Text14.SetText sFax
                    Reporte.Text16.SetText sRUC
                    Reporte.Text50.SetText sWeb
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                    Reporte.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                    Reporte.Text4.SetText xMontoTexto
                    Reporte.Text31.SetText xImpresionFE
                    
                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    Reporte.PrintOut False, 1, False, 1, 1
                    '----------------
                    
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub



Private Sub Form_Load()
   Dim nEmision As Integer
   Dim i As Integer
   Dim col As Integer
   Dim margentop As Integer
   Dim margenleft As Integer
      
   lImp1 = False
   lImp2 = False
   lImp3 = False
   
   If Calcular("select count(tItem) as Codigo from " & sTemporal & " where Estado='X' and nImpuesto1 > 0", Cn) > 0 Then
      Acumulado = Acumulado + nPorcentaje1
      lImp1 = True
   End If
   If Calcular("select count(tItem) as Codigo from " & sTemporal & " where Estado='X' and nImpuesto2 > 0", Cn) > 0 Then
      Acumulado = Acumulado + nPorcentaje2
      lImp2 = True
   End If
   If Calcular("select count(tItem) as Codigo from " & sTemporal & " where Estado='X' and nImpuesto3 > 0", Cn) > 0 Then
      Acumulado = Acumulado + nPorcentaje3
      lImp3 = True
   End If
   
   Acumulado = 1 + (Acumulado / 100)
   
   'Llena la Cabecera del Documento
   sCabecera = dbTemporal(sCaja, 29, "Item", "Int", _
                                     "tDocumento", "nVarChar(20)", _
                                     "tTipoDocumento", "nVarChar(2)", _
                                     "tTipoPedido", "nVarChar(2)", _
                                     "fFecha", "smalldatetime", _
                                     "tCodigoCliente", "nVarChar(5)", _
                                     "tMoneda", "nVarChar(3)", _
                                     "nNeto", "Float", _
                                     "nRecargo", "Float", _
                                     "nDescuento", "Float", _
                                     "nPrecioOficial", "Float", _
                                     "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "nPropina", "Float", _
                                     "nTotal", "Float", _
                                     "tCortesia", "nVarChar(4)", _
                                     "tEstadoDocumento", "nVarChar(2)", _
                                     "tTurno", "nVarChar(10)", _
                                     "tCaja", "nVarChar(3)", _
                                     "tObservacion", "nVarChar(250)", _
                                     "tUsuario", "nVarChar(15)", _
                                     "fRegistro", "smalldatetime", _
                                     "tEmision", "nVarChar(1)", _
                                     "tConsumo", "nVarChar(150)", "tautorizacion", "nVarChar(25)", "tCodigocontrol", "nVarChar(20)", "tDescuento", "nVarChar(3)")
   nCantidad = val(sTemp)
   
   sCodigoDescuento = Calcular("select tDescuento as Codigo From MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
   Isql = "select sum(nVenta - nImpuesto1 - nImpuesto2 - nImpuesto3) as nPrecioNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, " & _
          "sum(nImpuesto3) as nImpuesto3, sum(nVenta) + sum(case when vproducto.lley1=1 then " & sTemporal & ".ncantidad*(" & CDbl(sValorLey1) & ") else 0 end) as nVenta, sum(nDescuento*nCantidad) as nDescuento from " & sTemporal & " inner join vproducto on " & sTemporal & ".tcodigoproducto=vproducto.codigo where Estado='X'"
   
   Set RsTotales = Lib.OpenRecordset(Isql, Cn)
   txtDescuentoTotal.Caption = Format(RsTotales!nDescuento, "###,###,##0.00")
   lblTotal.Caption = Format(RsTotales!nVenta, "###,###,##0.00")
   txtMontoTotal.Caption = Format(RsTotales!nVenta, "###,###,##0.00")
     
   For i = 1 To nCantidad
       Isql = "insert into " & sCabecera & _
               "( item, fFecha, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, tDescuento) " & _
              " values ( " & i & ", " _
                           & "getdate(), " _
                           & RsTotales!nPrecioNeto / nCantidad & ", " _
                           & "0 , " _
                           & RsTotales!nDescuento / nCantidad & ", " _
                           & RsTotales!nVenta / nCantidad & ", " _
                           & RsTotales!nImpuesto1 / nCantidad & ", " _
                           & RsTotales!nImpuesto2 / nCantidad & ", " _
                           & RsTotales!nImpuesto3 / nCantidad & ", " _
                           & RsTotales!nVenta / nCantidad & ", '" _
                           & sCodigoDescuento & "') "
       Cn.Execute Isql
   Next
   
   Isql = "SELECT [" & sCabecera & "].tDocumento, [" & sCabecera & "].Item, [" & sCabecera & "].tCodigoCliente, TCLIENTE.tEmpresa AS Cliente, [" & sCabecera & "].tMoneda, [" & sCabecera & "].tEstadoDocumento, TCLIENTE.tIdentidad AS RUC, TCLIENTE.tDireccion AS Direccion, [" & sCabecera & "].nNeto, [" & sCabecera & "].nPrecioImpuesto1, [" & sCabecera & "].nPrecioImpuesto2, [" & sCabecera & "].fFecha, [" & sCabecera & "].nPrecioImpuesto3, [" & sCabecera & "].nPrecioOficial, [" & sCabecera & "].nRecargo, [" & sCabecera & "].nDescuento, [" & sCabecera & "].nVenta, [" & sCabecera & "].tTipoDocumento, [" & sCabecera & "].tUsuario, vTipoDocumento.Descripcion AS Documento, vCortesia.Descripcion AS Cortesia, [" & sCabecera & "].tConsumo, vMotivoDescuento.tResumido as Descuento , vmotivodescuento.codigo tdescuento " & _
          "FROM ((TCLIENTE RIGHT JOIN [" & sCabecera & "] ON TCLIENTE.tCodigoCliente = [" & sCabecera & "].tCodigoCliente) LEFT JOIN vTipoDocumento ON [" & sCabecera & "].tTipoDocumento = vTipoDocumento.Codigo) LEFT JOIN vCortesia ON [" & sCabecera & "].tCortesia = vCortesia.Codigo LEFT JOIN VMOTIVODESCUENTO ON [" & sCabecera & "].tdescuento = vMotivoDescuento.codigo "
   Set RsDocumento = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(9, grdDocumento, "No", 2, "Item", 400, 1, 0, "###", _
                                     "T/Documento", 2, "Documento", 1400, 0, 0, "", _
                                     "Cliente", 2, "Cliente", 3000, 0, 0, "", _
                                     "Neto", 2, "nNeto", 1000, 1, 0, "####,###,##0.00", _
                                     sImpuesto1, 2, "nPrecioImpuesto1", 850, 1, 0, "####,###,##0.00", _
                                     sImpuesto2, 2, "nPrecioImpuesto2", 850, 1, 0, "####,###,##0.00", _
                                     sImpuesto3, 2, "nPrecioImpuesto3", 850, 1, 0, "####,###,##0.00", _
                                     "Total", 2, "nVenta", 1200, 1, 0, "####,###,##0.00", _
                                     "Descto", 2, "nDescuento", 850, 1, 0, "####,###,##0.00")
   Set grdDocumento.DataSource = RsDocumento
   
   AsignaTotales
      
   'Emision x Impresora (Documentos) sin "Sin Documento"
'   If pais = "002" Then 'Ecuador
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' order by tTipoEmision", Cn)
'   Else
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 order by tTipoEmision", Cn)
'   End If

   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If
   
   
   nEmision = RsTipoDocumento.RecordCount
   If nEmision = 0 Then
      MsgBox "No se ha ingresado los Documentos por Caja", vbCritical, "Error Fatal"
      Unload Me
   End If
   
   ReDim mTipo(nEmision, 5)
       
   cmdTipoDocumento(0).Visible = False
   RsTipoDocumento.MoveFirst
    
   col = 0
   margentop = cmdTipoDocumento(0).Top
   margenleft = cmdTipoDocumento(0).Left
   wCortesia = False
   For i = 1 To nEmision
       If RsTipoDocumento!TTipoEmision = "00" Then
          wCortesia = True
       End If
       Load cmdTipoDocumento(i)
       If RsTipoDocumento!TTipoEmision <> "00" Then
          col = col + 1
          cmdTipoDocumento(i).Caption = RsTipoDocumento!Descripcion
          cmdTipoDocumento(i).Left = (col * 50) + ((col - 1) * cmdTipoDocumento(0).Width) + margenleft
          cmdTipoDocumento(i).Top = margentop
          cmdTipoDocumento(i).Visible = True
          
          mTipo(i, 1) = RsTipoDocumento!TTipoEmision
          mTipo(i, 2) = RsTipoDocumento!prefijo
          mTipo(i, 3) = RsTipoDocumento!Cliente
          mTipo(i, 4) = RsTipoDocumento!timpresora
          mTipo(i, 5) = RsTipoDocumento!Monto
       Else
          cmdTipoDocumento(i).Visible = False
       End If
       
       'Cantidad de Columnas
       If i Mod 4 = 0 Then
          col = 0
          margentop = margentop + cmdTipoDocumento(0).Height + 50
       End If
       RsTipoDocumento.MoveNext
   Next i
   ReDim mConsumo(nCantidad, 2)
   
   If sModulo = "PVCORP" Then
      Dim sXXX As String
      sXXX = Calcular("select tIdentidad as codigo from TCOMPANIA where tCodigoCliente ='" & sCliente & "'", Cn)
      sXXX = Calcular("select tCodigoCliente as Codigo from TCLIENTE where tIdentidad ='" & sXXX & "'", Cn)
      If sXXX <> "0" Then
         sCliente = sXXX
         Cn.Execute "update " & sCabecera & " set tCodigoCliente='" & sXXX & "'"
         txtRazonS.Caption = Calcular("select tEmpresa as Codigo from TCLIENTE where  tCodigoCliente='" & sXXX & "'", Cn)
         RsDocumento.Requery
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
   Dim nPos As Integer
   Select Case Index
          Case Is = 6  ' Monto Documento
                sTipo = ""
                frmNumPad.Show vbModal
                If wEnter Then
                   nPos = grdDocumento.Bookmark
  
                'validacionMontoMinimo
                   nMontoPedidoFacturarMInimo = 0
                   nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsDocumento!tTipoDocumento & "'", Cn)
                   If nMontoPedidoFacturarMInimo > 0 Then
                       If nMontoPedidoFacturarMInimo >= CDbl(sDescrip) Then
                           MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
                           wEnter = False
                           Exit Sub
                       End If
                   End If
                           

                   If nVeces = 0 Then
                      Isql = "Update " & sCabecera & " set nPrecioOficial = 0, nNeto = 0, nPrecioImpuesto1 = 0, nPrecioImpuesto2 = 0, nPrecioImpuesto3 = 0, nVenta = 0"
                      Cn.Execute Isql
                   End If
                   nVeces = nVeces + 1
                   txtMonto = Format(val(sDescrip), "###,###,##0.00")
                   
                   nTotal = val(sDescrip)
                   Select Case pais 'ok
                        Case "001" 'Bolivia
                            If lImp1 Then
                               xImpuesto1 = nTotal * nPorcentaje1 / 100
                            End If
                            If lImp2 Then
                               xImpuesto2 = nTotal * nPorcentaje2 / 100
                            End If
                            If lImp3 Then
                               xImpuesto3 = nTotal * nPorcentaje3 / 100
                            End If
                            
                        Case Else 'Peru, Ecuador
                            If lImp1 Then
                               xImpuesto1 = nTotal / Acumulado * nPorcentaje1 / 100
                            End If
                            If lImp2 Then
                               xImpuesto2 = nTotal / Acumulado * nPorcentaje2 / 100
                            End If
                            If lImp3 Then
                               xImpuesto3 = nTotal / Acumulado * nPorcentaje3 / 100
                            End If
                            
                    End Select
                                                                                                        
                   Isql = "Update " & sCabecera & " set nPrecioOficial = " & nTotal & ", " & _
                                                  "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                  "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                  "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                  "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                  "nVenta = " & nTotal & " " & _
                                                  "where item = " & nPos
                                                  
                   Cn.Execute Isql
                   RsDocumento.Requery
                   
                   
                   AsignaTotales
                   txtDiferencia.Caption = Format(val(str(txtMontoTotal.Caption)) - val(Calcular("select sum(nVenta) as Codigo From " & sCabecera, Cn)), "###,###,##0.00")
                   grdDocumento.Bookmark = nPos
                End If
               
               
          Case Is = 7 ' Limpiar
                nVeces = 1
                Isql = "Update " & sCabecera & " set nPrecioOficial = 0, nNeto = 0, nPrecioImpuesto1 = 0, nPrecioImpuesto2 = 0, nPrecioImpuesto3 = 0, nVenta = 0"
                Cn.Execute Isql
                CmdOpcion(6).Enabled = True
                RsDocumento.Requery
                RsDocumento.MoveFirst
                             
          Case Is = 10 ' Primero
               MoverPuntero Primero, grdDocumento
               'AsignaDetalle
               AsignaTotales

          Case Is = 11 ' Previo
               MoverPuntero previo, grdDocumento
               'AsignaDetalle
               AsignaTotales

          Case Is = 12 ' Siguiente
               MoverPuntero siguiente, grdDocumento
               'AsignaDetalle
               AsignaTotales

          Case Is = 13 ' Ultimo
               MoverPuntero Ultimo, grdDocumento
               'AsignaDetalle
               AsignaTotales
                    
          Case Is = 15 ' Salir
               wEnter = False
               Unload Me

          Case Is = 17 ' Cliente
               sTemp = ""
               Isql = "select * from vCliente where lActivo=1 order by Descripcion"
              Isql = "exec usp_Inforest_ObtieneClientesFactura '','" & xTipoDocSeleccionado & "'," & CDbl(txtMonto.Caption) & ", '" & sCaja & "'"

               frmBusquedaRapida.CmdOpcion(1).Enabled = True
               frmBusquedaRapida.CmdOpcion(2).Enabled = True
               frmBusquedaRapida.CmdOpcion(3).Enabled = True
               frmBusquedaRapida.nPredeterm = 1
               'imprimeDNI
               Select Case pais
                Case "001" 'Bolivia
                    Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                               "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
               
               End Select
               If sVar1 = "02" Then
                  sTemp = Calcular("SELECT dbo.TCLIENTE.tIdentidad AS codigo FROM dbo.TDELIVERY INNER JOIN dbo.TCLIENTE ON dbo.TDELIVERY.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente WHERE tCodigoDelivery ='" & sCliente & "'", Cn)
                  sTemp = IIf(sTemp = "0", "", sTemp)
               Else
                  sTemp = ""
               End If
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True Then
                  If sVar1 = "02" Then
                     Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sCodigo & "' where tCodigoDelivery ='" & sCliente & "'"
                  End If
                  
                  nPos = grdDocumento.Bookmark
                  Isql = "Update " & sCabecera & " set tCodigoCliente='" & sCodigo & "' where item = " & nPos
                  Cn.Execute Isql
                  txtRazonS.Caption = sDescrip
                  RsDocumento.Requery
                  
                  grdDocumento.Bookmark = nPos
                  'AsignaDetalle
               Else
                  nPos = grdDocumento.Bookmark
                  Isql = "Update " & sCabecera & " set tCodigoCliente='' where item = " & nPos
                  Cn.Execute Isql
                  txtRazonS.Caption = ""
                  RsDocumento.Requery
                  grdDocumento.Bookmark = nPos
                  'AsignaDetalle
               End If
               
          Case Is = 18 ' Cortesia
               If Not wCortesia Then
                  MsgBox "Caja sin Cortesía", vbCritical, sMensaje
                  Exit Sub
               End If
          
               If Supervisor("04") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If

               sTemp = ""
               Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True Then
                  nPos = grdDocumento.Bookmark
                  Isql = "Update " & sCabecera & " set tTipoDocumento='00', " & _
                         "tCodigoCliente='', " & _
                         "tCortesia= '" & sCodigo & "' " & _
                         " where item = " & nPos

                  Cn.Execute Isql
                  RsDocumento.Requery
                  grdDocumento.Bookmark = nPos
                  'AsignaDetalle
               End If
          
   End Select
End Sub

Public Sub cmdTipoDocumento_Click(Index As Integer)
   Dim nRespuesta As Integer
   Dim xConsumo As String
   RsTipoDocumento.MoveFirst
   RsTipoDocumento.Move (Index - 1)
   
        'validacionMontoMinimo
           nMontoPedidoFacturarMInimo = 0
           nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
           If nMontoPedidoFacturarMInimo > 0 Then
               If nMontoPedidoFacturarMInimo >= CDbl(txtMonto.Caption) Then
                   MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
                   wEnter = False
                   Exit Sub
               End If
           End If
          
   
        'validacionMontoMaximo
        nMontoPedidoFacturar = 0
        nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
        If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar < CDbl(txtMonto.Caption) Then
            MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
            wEnter = False
            Exit Sub
        End If
        End If

        '----VALIDA BLOQUEO DE DOCUMENTOS POR CONSUMO CON BONIFICACIONES / 03012020 / GCAA -------------------
        If lbloqueadocconsumo = True And lConsumo2 = True Then
            Dim n1 As Double
            Dim ix As Integer
                For ix = 0 To frmDocumento.grdDetalle2.ApproxCount - 1
                    frmDocumento.grdDetalle2.Row = ix
                    n1 = CDbl(frmDocumento.grdDetalle2.Columns(3))
                   If n1 = 0 Then
                        MsgBox "Información: No se puede emitir el documento por consumo con bonificaciones (item con precio 0)", vbInformation, sMensaje
                        Exit Sub
                   End If
                Next
        
        End If
        
        If lBloqueoAfectoInafecto And xConsumo <> "" Then
            Dim inafecto As Integer
            Dim afecto As Integer
            inafecto = Calcular("select count(*) as codigo from " & sTemporal & " where Estado='X' and round(nimpuesto1,2)=0 ", Cn)
            afecto = Calcular("select count(*) as codigo from " & sTemporal & " where Estado='X' and round(nimpuesto1,2)>0 ", Cn)
            
            If inafecto > 0 And afecto > 0 Then
                MsgBox "Bloqueo de emision: Documento por consumo con afectos e inafectos", vbInformation, sMensaje
                'RsDetalle.Requery
                Exit Sub
            End If
            
        End If



   If lConsumo2 = True Then
      xConsumo = ""
      nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
      If nRespuesta = vbYes Then
         frmKeyBoard.txtResultado = tTextoConsumo
         frmKeyBoard.Show vbModal
         If sDescrip = "" Or Not wEnter Then
            MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
            Exit Sub
         End If
         xConsumo = sDescrip
      ElseIf nRespuesta = vbCancel Then
         Exit Sub
      End If
   End If
   
   nPos = grdDocumento.Bookmark
   If sModulo = "PVCORP" Then
      Isql = "Update " & sCabecera & " set tTipoDocumento='" & mTipo(Index, 1) & "', tCortesia='' where item = " & nPos
   Else
      Isql = "Update " & sCabecera & " set tTipoDocumento='" & mTipo(Index, 1) & "', tCodigoCliente='', tCortesia='', tConsumo ='" & xConsumo & "' where item = " & nPos
   End If
   Cn.Execute Isql
   xTipoDocSeleccionado = mTipo(Index, 1)
   RsDocumento.Requery
   grdDocumento.Bookmark = nPos
   If mTipo(Index, 3) And (mTipo(Index, 5) = 0 Or val(mTipo(Index, 5)) < RsDocumento!nVenta) Then
      CmdOpcion_Click (17)
   End If
End Sub

Public Sub AsignaTotales()
    If RsDocumento.EOF Or RsDocumento.BOF Then
     RsDocumento.MoveFirst
    End If
    
   nNeto = IIf(IsNull(RsDocumento!nNeto), 0, RsDocumento!nNeto)
   nPrecioOficial = IIf(IsNull(RsDocumento!nPrecioOficial), 0, RsDocumento!nPrecioOficial)
   nTotal = IIf(IsNull(RsDocumento!nVenta), "0", RsDocumento!nVenta)
   txtNeto.Caption = Format(nNeto, "###,###,##0.00")
   txtDImporte.Caption = Format(RsDocumento!nDescuento, "###,###,##0.00")
   txtTotal.Caption = Format(nTotal, "###,###,##0.00")
   txtDescuento.Caption = IIf(IsNull(RsDocumento!DESCUENTO), "", RsDocumento!DESCUENTO)
   txtMonto = Format(nTotal, "###,###,##0.00")
   txtRazonS.Caption = IIf(IsNull(RsDocumento!Cliente), "", RsDocumento!Cliente)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sCabecera
   Set RsDocumento = Nothing
   Set RsTotales = Nothing
   Set RsTipoDocumento = Nothing
   Set RsArea = Nothing
   Set RsDetalle = Nothing
   Set RsSuma = Nothing
   Set frmGeneracion = Nothing
End Sub

Private Sub grdDocumento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignaTotales
End Sub

Public Sub Facturar(RsTD As Recordset, TipoDocumento As String, tipo As String)
   Dim RsGenera As Recordset
   Dim RsDeta As Recordset
   Dim RsSuma As Recordset
   
   Dim nCont As Double
   Dim sGenera As String
   Dim sDetalle As String
   Dim sTipoEmision As String
   Dim sSerie As String
   Dim sCorrela As String
   Dim sPrefijo As String
   Dim nPos As Variant
   
   Dim mDocumento() As String
   Dim mTipoEmision() As String
   Dim mTipoDocumento() As String
   
   Dim nDocumento As Integer
   Dim sDetalleConsumo As String
   Dim sCortesia As String
   Dim i As Integer
      
   'Crea el Temporal del detalle a Facturar DPEDIDO
   sDetalle = dbTemporal(sCaja, 17, "tDocumento", "nVarChar(20)", _
                                    "tItem", "nVarChar(3)", _
                                    "tCodigoPedido", "nVarChar(10)", _
                                    "tCodigoProducto", "nVarChar(7)", _
                                    "nPrecioNeto", "Float", _
                                    "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                    "nPrecioVenta", "Float", _
                                    "nPrecioOficial", "Float", _
                                    "nRecargo", "Float", _
                                    "nDescuento", "Float", _
                                    "nCantidad", "Float", _
                                    "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                    "nVenta", "Float")
                                    
   'Inserta el DPEDIDO Escogido a Facturar
   Select Case tipo
          Case Is = "Rapido"
               Isql = "Insert into " & sDetalle & _
                      "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                      "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
                      "where tCodigoPedido ='" & sPedido & "' and tEstadoItem = 'N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0)"
                      
          Case Is = "Division"
               Isql = "Insert into " & sDetalle & _
                      "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                      "select  tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sTemporal
   End Select
   Cn.Execute Isql
  
   sGenera = dbTemporal(sCaja, 6, "tItem", "nVarChar(3)", _
                                  "tCodigoProducto", "nVarChar(7)", _
                                  "nPrecioVenta", "Float", _
                                  "nCantidad", "Float", _
                                  "nVenta", "Float", _
                                  "tDocumento", "nVarChar(20)")
                                    
   'Genera el Temporal a Facturar
   Set RsGenera = Lib.OpenRecordset("select * from " & sGenera, Cn)
   Set RsDeta = Lib.OpenRecordset("select * from " & sDetalle, Cn)
   RsDeta.MoveFirst
   Do While Not RsDeta.EOF
      RsGenera.Filter = ("tCodigoProducto ='" & RsDeta!tCodigoProducto & "'")
      If Not RsGenera.EOF Then
         RsGenera.MoveFirst
      End If
      RsGenera.Find "nPrecioVenta =" & RsDeta!nprecioVenta
      If RsGenera.EOF Then
         RsGenera.AddNew
         RsGenera!tCodigoProducto = RsDeta!tCodigoProducto
         RsGenera!nprecioVenta = RsDeta!nprecioVenta
         RsGenera!nCantidad = RsDeta!nCantidad
         RsGenera!nVenta = RsDeta!nVenta
      Else
         RsGenera!nCantidad = RsGenera!nCantidad + RsDeta!nCantidad
         RsGenera!nVenta = RsGenera!nVenta + RsDeta!nVenta
      End If
      RsGenera.Update
      RsGenera.Requery
      RsDeta.MoveNext
   Loop
   RsGenera.Filter = adFilterNone
     
   'Chequea el Tipo de Documento
   'Factura / Botela / Cortesia
   'Variable / Ticket
   RsTD.Requery
   RsTD.MoveFirst
   RsTD.Find ("tTipoEmision='" & TipoDocumento & "'")
   If RsTD.EOF Then
      MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
      Exit Sub
   End If

   'Tipo de Emision
   sTipoEmision = ""
   If sTipoEmision = "Consolidado" And RsTD!tFormulario = "02" Then
      
      'Genero el Documento
      nDocumento = 1
      sSerie = RsTD!tSerie
      sCorrela = Lib.Correlativo(RsTD!tUltimoNumero, 9)
      sPrefijo = RsTD!prefijo
      sDocumento = sPrefijo & sSerie & sCorrela
      ReDim Preserve mDocumento(nDocumento)
      ReDim Preserve mTipoEmision(nDocumento)
      ReDim Preserve mTipoDocumento(nDocumento)
      
      mDocumento(nDocumento) = sDocumento
      mTipoEmision(nDocumento) = sTipoEmision
      mTipoDocumento(nDocumento) = RsTD!TTipoEmision
      
      nPos = RsTD.Bookmark
      Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & mTipoDocumento(nDocumento) & "' and tCaja ='" & sCaja & "'"
      RsTD.Requery
      RsTD.Bookmark = nPos
                 
      RsGenera.MoveFirst
      nCont = 0
      Do While Not RsGenera.EOF
      
         If nCont = nFItem Then
            'Genero el Documento
            nCont = 1
            nDocumento = nDocumento + 1
            sSerie = RsTD!tSerie
            sCorrela = Lib.Correlativo(RsTD!tUltimoNumero, 9)
            sPrefijo = RsTD!prefijo
            sDocumento = sPrefijo & sSerie & sCorrela
            
            ReDim Preserve mDocumento(nDocumento)
            ReDim Preserve mTipoEmision(nDocumento)
            ReDim Preserve mTipoDocumento(nDocumento)
            mDocumento(nDocumento) = sDocumento
            mTipoEmision(nDocumento) = sTipoEmision
            mTipoDocumento(nDocumento) = RsTD!TTipoEmision
            
            nPos = RsTD.Bookmark
            Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & mTipoDocumento(nDocumento) & "' and tCaja ='" & sCaja & "'"
            RsTD.Requery
            RsTD.Bookmark = nPos
            
            RsGenera!tDocumento = sDocumento
            RsGenera!tItem = Llena(nCont, 3)
            RsGenera.Update
                        
            RsDeta.Filter = ("tCodigoProducto ='" & RsGenera!tCodigoProducto & "'")
            RsDeta.MoveFirst
            Do While Not RsDeta.EOF
               If RsDeta!nprecioVenta = RsGenera!nprecioVenta Then
                  Isql = "Update " & sDetalle & " set tDocumento = '" & sDocumento & "' " & _
                         "where tCodigoPedido = '" & RsDeta!tCodigoPedido & "' and tItem = '" & RsDeta!tItem & "'"
                  Cn.Execute Isql
               End If
               RsDeta.MoveNext
            Loop
         Else
            nCont = nCont + 1
            RsGenera!tDocumento = sDocumento
            RsGenera!tItem = Llena(nCont, 3)
            RsGenera.Update
                       
            RsDeta.Filter = ("tCodigoProducto ='" & RsGenera!tCodigoProducto & "'")
            RsDeta.MoveFirst
            Do While Not RsDeta.EOF
               If RsDeta!nprecioVenta = RsGenera!nprecioVenta Then
                  Isql = "Update " & sDetalle & " set tDocumento = '" & sDocumento & "' " & _
                         "where tCodigoPedido = '" & RsDeta!tCodigoPedido & "' and tItem = '" & RsDeta!tItem & "'"
                  Cn.Execute Isql
               End If
               RsDeta.MoveNext
            Loop
            
         End If
         RsGenera.MoveNext
         If RsGenera.EOF Then
            Exit Do
         End If
      Loop
      RsGenera.Requery
      RsDeta.Filter = adFilterNone
      RsDeta.Requery
      
      'Genero la Cabecera del MDOCUMENTO
      For i = 1 To nDocumento
          'Calcula el total de la cabecera
          Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                         " from " & sDetalle & " where tDocumento ='" & mDocumento(i) & "' group by tDocumento ", Cn)
            
            Select Case pais 'ok
                Case "001" 'Bolivia
                    tcodigoControl = devuelveCodigoControl(sCaja, Mid(mDocumento(i), 7, 9), tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                    
                Case Else 'Peru, Ecuador
                    tAutorizacion = ""
                    tcodigoControl = ""
                    
            End Select
                      
          'Inserta la Cabecera del Documento
          Isql = "Insert into MDOCUMENTO " & _
                 " ( tDocumento, tTipoDocumento, fFecha, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, fRegistro, " & _
                 "nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta,tautorizacion,tcodigocontrol, fdiacontable , lImpresionMonedaExtranjera) " & _
                 "Values(   '" & mDocumento(i) & "', " _
                     & "'" & mTipoDocumento(i) & "', " _
                     & "{fn curdate()}, " _
                     & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                     & "'01', " _
                     & "'" & sCaja & "', " _
                     & "'" & sSalon & "', " _
                     & "'" & sTurno & "', " _
                     & "'" & Mid(sUsuario, 1, 15) & "', " _
                     & " getdate() , " & RsSuma!nNeto & ", " _
                     & "0, 0, 0, " _
                     & RsSuma!nImpuesto1 & ", " _
                     & RsSuma!nImpuesto2 & ", " _
                     & RsSuma!nImpuesto3 & ", " _
                     & RsSuma!nVenta & ",'" & tAutorizacion & "','" & tcodigoControl & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & ")"
          Cn.Execute Isql
          
         'Actualiza el DPEDIDO con el Temporal sDetalle
         RsDeta.MoveFirst
         Do While Not RsDeta.EOF
            If RsDeta!tDocumento = mDocumento(i) Then
               Isql = "Update DPEDIDO set tDocumento = '" & mDocumento(i) & "' , " & _
                                        " tFacturado = 'F' " & _
                                        " where tCodigoPedido = '" & RsDeta!tCodigoPedido & "' and tCodigoProducto ='" & RsDeta!tCodigoProducto & "' and tItem='" & RsDeta!tItem & "'"
                                        
               Cn.Execute Isql
               
               'PARA NO FISCALES
               Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & mTipoDocumento(i) & "') where  tCodigoPedido ='" & RsDeta!tCodigoPedido & "' and tDocumento ='" & mDocumento(i) & "'"
        
            End If
            RsDeta.MoveNext
         Loop
         
         'Inserta Detalle de Documento
         Isql = "Insert into DDOCUMENTO " & _
                "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                "select  '" & mDocumento(i) & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sDetalle & _
                " where tDocumento ='" & mDocumento(i) & "'"
         Cn.Execute Isql
      Next i
      
   Else     'Por Consumo  / Detallado
      
      'Genero el Documento
      nDocumento = 1
      sSerie = RsTD!tSerie
      sCorrela = Lib.Correlativo(RsTD!tUltimoNumero, 9)
      sPrefijo = RsTD!prefijo
      sDocumento = sPrefijo & sSerie & sCorrela
      ReDim Preserve mDocumento(nDocumento)
      ReDim Preserve mTipoEmision(nDocumento)
      ReDim Preserve mTipoDocumento(nDocumento)
      
      mDocumento(nDocumento) = sDocumento
      mTipoEmision(nDocumento) = sTipoEmision
      mTipoDocumento(nDocumento) = RsTD!TTipoEmision
      
      nPos = RsTD.Bookmark
      Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & mTipoDocumento(nDocumento) & "' and tCaja ='" & sCaja & "'"
      RsTD.Requery
      RsTD.Bookmark = nPos
                 
      RsGenera.MoveFirst
      nCont = 0
      Do While Not RsGenera.EOF
      
         nCont = nCont + 1
         RsGenera!tDocumento = sDocumento
         RsGenera!tItem = Llena(nCont, 3)
         RsGenera.Update
                    
         RsDeta.Filter = ("tCodigoProducto ='" & RsGenera!tCodigoProducto & "'")
         RsDeta.MoveFirst
         Do While Not RsDeta.EOF
            If RsDeta!nprecioVenta = RsGenera!nprecioVenta Then
               Isql = "Update " & sDetalle & " set tDocumento = '" & sDocumento & "' " & _
                      "where tCodigoPedido = '" & RsDeta!tCodigoPedido & "' and tItem = '" & RsDeta!tItem & "'"
               Cn.Execute Isql
            End If
            RsDeta.MoveNext
         Loop
            
         RsGenera.MoveNext
         If RsGenera.EOF Then
            Exit Do
         End If
      Loop
            
      RsGenera.Requery
      RsDeta.Filter = adFilterNone
      RsDeta.Requery
      
      'Genero la Cabecera del MDOCUMENTO
      For i = 1 To nDocumento
          'Calcula el total de la cabecera
          Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                         " from " & sDetalle & " where tDocumento ='" & mDocumento(i) & "' group by tDocumento ", Cn)
            
            Select Case pais ' ok
                    Case "001" 'Bolivia
                        tcodigoControl = devuelveCodigoControl(sCaja, Mid(mDocumento(i), 7, 9), tAutorizacion, tDosificacion, sCliente, RsSuma!nVenta)
                    Case Else 'Peru, Ecuador
                        tAutorizacion = ""
                        tcodigoControl = ""
            End Select
            
          'Inserta la Cabecera del Documento
          Isql = "Insert into MDOCUMENTO " & _
                 " ( tDocumento, tTipoDocumento, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, fRegistro, " & _
                 "nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta,tautorizacion,tcodigocontrol, fdiacontable, lImpresionMonedaExtranjera ) " & _
                 "Values(   '" & mDocumento(i) & "', " _
                     & "'" & mTipoDocumento(i) & "', " _
                     & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                     & "'01', " _
                     & "'" & sCaja & "', " _
                     & "'" & sSalon & "', " _
                     & "'" & sTurno & "', " _
                     & "'" & Mid(sUsuario, 1, 15) & "', " _
                     & " getdate() , " & RsSuma!nNeto & ", " _
                     & "0, 0, 0, " _
                     & RsSuma!nImpuesto1 & ", " _
                     & RsSuma!nImpuesto2 & ", " _
                     & RsSuma!nImpuesto3 & ", " _
                     & RsSuma!nVenta & ",'" & tAutorizacion & "' ,'" & tcodigoControl & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " )"
          Cn.Execute Isql
          
          nCargo = RsSuma!nVenta
          
         'Actualiza el DPEDIDO con el Temporal sDetalle
         RsDeta.MoveFirst
         Do While Not RsDeta.EOF
            If RsDeta!tDocumento = mDocumento(i) Then
               Isql = "Update DPEDIDO set tDocumento = '" & mDocumento(i) & "' , " & _
                                        " tFacturado = 'F' " & _
                                        " where tCodigoPedido = '" & RsDeta!tCodigoPedido & "' and tCodigoProducto ='" & RsDeta!tCodigoProducto & "' and tItem='" & RsDeta!tItem & "'"
               Cn.Execute Isql
               'PARA NO FISCALES
               Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & mTipoDocumento(i) & "') where  tCodigoPedido ='" & RsDeta!tCodigoPedido & "' and tDocumento ='" & mDocumento(i) & "'"
            End If
            RsDeta.MoveNext
         Loop
         
         'Inserta Detalle de Documento
         Isql = "Insert into DDOCUMENTO " & _
                "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                "select  '" & mDocumento(i) & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From " & sDetalle & _
                " where tDocumento ='" & mDocumento(i) & "'"
         Cn.Execute Isql
      Next i
   End If
      
   'Imprime Documentos
   For i = 1 To nDocumento
     Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.fRegistro, MIN(dbo.MPEDIDO.tObservacion) AS tObservacion, " & _
              "MIN(dbo.MPEDIDO.tCodigoPedido) AS tCodigoPedido, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tCodigoProducto, dbo.TPRODUCTO.tDetallado AS Producto, MIN(dbo.DDOCUMENTO.tItem) AS tItem, SUM(dbo.DDOCUMENTO.nCantidad) AS nCantidad, SUM(dbo.DDOCUMENTO.nVenta) AS nPrecioVenta" & _
              "FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN dbo.MPEDIDO ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
              "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.TCLIENTE.tEmpresa, dbo.TCLIENTE.tIdentidad, dbo.TCLIENTE.tDireccion, dbo.DDOCUMENTO.tCodigoProducto , dbo.mDocumento.tCaja, dbo.mDocumento.tUsuario, dbo.TPRODUCTO.tDetallado " & _
              "HAVING (dbo.MDOCUMENTO.tDocumento = N'" & mDocumento(i) & "') order by tItem"
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)

       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
       
        'SUNAT
        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, RsTipoDocumento!timpresora)
        codigoImpresora = RsTipoDocumento!timpresora
        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & mDocumento(i) & "' "

          'Configura la Impresora
          Imprimir (RsTipoDocumento!timpresora)
          
          Printer.FontName = sFont
          Printer.FontBold = False
          If RsTipoDocumento!TTipoEmision = "01" Then
             ImprimeFacturaN RsImpresion, sEmpresa, TipoDocumento
             NFactura = sCorrela
          Else
             ImprimeBoletaN RsImpresion, sEmpresa, TipoDocumento
           End If
       End If
     
'       If pais = "002" Then
'          Dim sXML As String
'          sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'          GeneraFacturaElectronica sXML, sDocumento
'       End If
       
       If pais = "002" And lFEEcuador = False Then
       Dim sXML As String
          sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
          GeneraFacturaElectronica sXML, sDocumento
       End If
       
       If lFEEcuador Then
        If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
            MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
        End If
       End If
       
       Set RsImpresion = Nothing
   Next i
End Sub


'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
                fImpresionDiaContable = obtieneDiaContable
End Function
'diaContable

Private Sub CalculaAplicaTope(nTope As Double)
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    Dim nPVenta As Double
    Dim RsOferta As Recordset
    Dim xDescuento As Double
    Dim nImpuesto1 As Double
    Dim nImpuesto2 As Double
    Dim nImpuesto3 As Double
    Dim nPBase As Double
    Dim RsDetalle As Recordset
    
    Set RsDetalle = Lib.OpenRecordset("SELECT * FROM " & sTemporal & " LEFT OUTER JOIN dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where Estado='X'", Cn)
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & sPedido & "' and lDescuento=1", Cn)

If RsDetalle.RecordCount <> 0 Then
   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
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
         
         xDescuento = (RsDetalle!nPrecioOficial - nOferta) * (RsDetalle!nCantidad * 100 / nSuma)
         nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((nTope * xDescuento / 100) / RsDetalle!nCantidad)
         
          Select Case pais ' ok
            Case "001" 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
            Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
          End Select
         Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = 0, " & _
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
                "tCortesia = '' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If
End Sub

Private Sub RevierteDocumento(ByVal Documento As String, ByVal Pedido As String, ByVal Mesa As String, ByVal Caja As String)
    Dim UltimoCorrelativo As String
    Dim TipoDocumento As String

    RsSuma.Requery
    RsSuma.MoveFirst
    Do While Not RsSuma.EOF
    
        TipoDocumento = Calcular("select tTipoDocumento as Codigo from mDocumento where tDocumento = '" & RsSuma!tDocumento & "'", Cn)
        
        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & Mesa & "'"
        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsSuma!tDocumento & "'"
        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsSuma!tDocumento & "'"
        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsSuma!tDocumento & "'"
        
        UltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & Caja & "' and tTipoDocumento='" & TipoDocumento & "'", Cn)
        UltimoCorrelativo = Right(UltimoCorrelativo, 9)
        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & UltimoCorrelativo & "' where tTipoEmision ='" & TipoDocumento & "' and tCaja ='" & Caja & "'"
        
        Cn.Execute "Update DPEDIDO set lregistroventa = NULL, tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & Pedido & "' "
        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & Pedido & "'"
    RsSuma.MoveNext
    Loop
    
End Sub


