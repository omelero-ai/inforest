VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGuiasDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Guias de Remisión"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmGuiasDetalle.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11880
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
      Height          =   375
      Index           =   2
      Left            =   11205
      Picture         =   "frmGuiasDetalle.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   416
      Width           =   585
   End
   Begin VB.PictureBox PicNavegacion 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11820
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   7665
      Width           =   11880
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   480
         Picture         =   "frmGuiasDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   960
         Picture         =   "frmGuiasDetalle.frx":0A86
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   0
         Picture         =   "frmGuiasDetalle.frx":0FC8
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   5
         Left            =   11265
         Picture         =   "frmGuiasDetalle.frx":150A
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   4
         Left            =   10785
         Picture         =   "frmGuiasDetalle.frx":1A4C
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   10305
         Picture         =   "frmGuiasDetalle.frx":1F8E
         Style           =   1  'Graphical
         TabIndex        =   51
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
         Height          =   225
         Left            =   4680
         TabIndex        =   57
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11820
      TabIndex        =   43
      Top             =   6930
      Width           =   11880
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
         Index           =   5
         Left            =   5220
         Picture         =   "frmGuiasDetalle.frx":24D0
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   30
         Width           =   1245
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
         Index           =   1
         Left            =   0
         Picture         =   "frmGuiasDetalle.frx":25D2
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   30
         Width           =   1245
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
         Index           =   2
         Left            =   1305
         Picture         =   "frmGuiasDetalle.frx":2B04
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   30
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Anular"
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
         Left            =   3915
         Picture         =   "frmGuiasDetalle.frx":3036
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   30
         Width           =   1245
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
         Index           =   3
         Left            =   2610
         Picture         =   "frmGuiasDetalle.frx":3138
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   30
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Imprimir"
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
         Left            =   6510
         Picture         =   "frmGuiasDetalle.frx":323A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   30
         Width           =   1245
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
         Index           =   0
         Left            =   7815
         Picture         =   "frmGuiasDetalle.frx":37C4
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   30
         Width           =   1245
      End
   End
   Begin VB.TextBox txtDomicilioPartida 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8325
      TabIndex        =   8
      Top             =   1173
      Width           =   3480
   End
   Begin VB.TextBox txtDomicilioDestino 
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
      Height          =   315
      Left            =   8325
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   817
      Width           =   3480
   End
   Begin VB.Frame fraObservacion 
      Caption         =   " Observación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2295
      TabIndex        =   34
      Top             =   5940
      Width           =   9525
      Begin VB.TextBox txtObservacion 
         Height          =   675
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   9300
      End
   End
   Begin VB.TextBox txtTotal 
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
      Height          =   315
      Left            =   1035
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   6615
      Width           =   1185
   End
   Begin VB.TextBox txtTara 
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
      Height          =   315
      Left            =   1035
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   6277
      Width           =   1185
   End
   Begin VB.TextBox txtNeto 
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
      Height          =   315
      Left            =   1035
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   5940
      Width           =   1185
   End
   Begin VB.TextBox txtCorrelativo 
      BackColor       =   &H00F0EADB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "000000000"
      Top             =   75
      Width           =   1140
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H00F0EADB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "00000"
      Top             =   75
      Width           =   780
   End
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
      Height          =   375
      Index           =   0
      Left            =   6570
      Picture         =   "frmGuiasDetalle.frx":38B6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   416
      Width           =   585
   End
   Begin VB.TextBox txtDocumento 
      BackColor       =   &H00F0EADB&
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
      Height          =   315
      Left            =   4365
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   446
      Width           =   2130
   End
   Begin VB.TextBox txtDestinatario 
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
      Height          =   315
      Left            =   8325
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   446
      Width           =   2805
   End
   Begin VB.TextBox txtTransportista 
      BackColor       =   &H00F0EADB&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4365
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   817
      Width           =   2130
   End
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
      Height          =   375
      Index           =   1
      Left            =   6570
      Picture         =   "frmGuiasDetalle.frx":39B8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   787
      Width           =   585
   End
   Begin VB.Frame fraDetalle 
      Caption         =   " Detalle de los Productos "
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
      Height          =   4365
      Left            =   0
      TabIndex        =   9
      Top             =   1530
      Width           =   11805
      Begin VB.CommandButton cmdAgproducto 
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
         Height          =   645
         Left            =   5265
         TabIndex        =   63
         Top             =   2925
         Width           =   1275
      End
      Begin VB.CommandButton cmdEliproducto 
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
         Height          =   645
         Left            =   5265
         TabIndex        =   62
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Height          =   540
         Left            =   6600
         TabIndex        =   17
         Top             =   3510
         Width           =   5115
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total del Documento :"
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
            Left            =   2280
            TabIndex        =   19
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label txtTotal1 
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
            Left            =   3720
            TabIndex        =   18
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   60
         TabIndex        =   14
         Top             =   3510
         Width           =   5115
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Total de la Guía :"
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
            Left            =   2595
            TabIndex        =   16
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label txtTotal2 
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
            Left            =   3720
            TabIndex        =   15
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   0
         Left            =   5250
         Picture         =   "frmGuiasDetalle.frx":3ABA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   2
         Left            =   5250
         Picture         =   "frmGuiasDetalle.frx":4044
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   885
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   1
         Left            =   5250
         Picture         =   "frmGuiasDetalle.frx":45CE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2175
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   3
         Left            =   5250
         Picture         =   "frmGuiasDetalle.frx":4910
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1530
         Width           =   1275
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle2 
         Height          =   3270
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5768
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
         Caption         =   "Productos que van a generar una Guía"
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
      Begin TrueOleDBGrid80.TDBGrid grdDetalle1 
         Height          =   3270
         Left            =   6600
         TabIndex        =   21
         Top             =   240
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5768
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
         Caption         =   "Productos por generar una Guía"
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
   Begin MSDataListLib.DataCombo cboMotivo 
      Height          =   315
      Left            =   1215
      TabIndex        =   3
      Top             =   810
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   15788763
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   345
      Left            =   1215
      TabIndex        =   2
      Top             =   435
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109772801
      CurrentDate     =   37539.2083333333
   End
   Begin MSDataListLib.DataCombo cboChofer 
      Height          =   315
      Left            =   4365
      TabIndex        =   7
      Top             =   1170
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   15788763
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
   Begin MSDataListLib.DataCombo cboVehiculo 
      Height          =   315
      Left            =   1215
      TabIndex        =   4
      Top             =   1170
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   15788763
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
   Begin MSDataListLib.DataCombo cboEstado 
      Height          =   315
      Left            =   4380
      TabIndex        =   61
      Top             =   90
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   15788763
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
      Caption         =   "Vehículo :"
      Height          =   195
      Index           =   5
      Left            =   450
      TabIndex        =   59
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Chofer :"
      Height          =   195
      Index           =   9
      Left            =   3765
      TabIndex        =   58
      Top             =   1230
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Estado :"
      Height          =   195
      Index           =   6
      Left            =   3735
      TabIndex        =   42
      Top             =   135
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Peso Neto :"
      Height          =   195
      Index           =   13
      Left            =   105
      TabIndex        =   41
      Top             =   6000
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Peso Bruto :"
      Height          =   195
      Index           =   15
      Left            =   75
      TabIndex        =   40
      Top             =   6675
      Width           =   870
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Tara :"
      Height          =   195
      Index           =   14
      Left            =   525
      TabIndex        =   39
      Top             =   6337
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Dom. Partida :"
      Height          =   195
      Index           =   12
      Left            =   7275
      TabIndex        =   38
      Top             =   1230
      Width           =   1005
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Dom. Destino :"
      Height          =   195
      Index           =   11
      Left            =   7230
      TabIndex        =   36
      Top             =   870
      Width           =   1050
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Destinatario :"
      Height          =   195
      Index           =   10
      Left            =   7350
      TabIndex        =   30
      Top             =   510
      Width           =   930
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Emisión :"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   29
      Top             =   510
      Width           =   1125
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Guía :"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   28
      Top             =   135
      Width           =   1020
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Motivo :"
      Height          =   195
      Index           =   2
      Left            =   615
      TabIndex        =   27
      Top             =   870
      Width           =   570
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Transportista :"
      Height          =   195
      Index           =   8
      Left            =   3315
      TabIndex        =   23
      Top             =   870
      Width           =   1005
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Documento :"
      Height          =   195
      Index           =   7
      Left            =   3405
      TabIndex        =   22
      Top             =   510
      Width           =   915
   End
End
Attribute VB_Name = "frmGuiasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDetalle1 As Recordset
Dim RsDetalle2 As Recordset
Dim RsMotivo As Recordset
Dim RsChofer As Recordset
Dim RsVehiculo As Recordset
Dim RsEstado As Recordset
Dim RsImpresion As Recordset
Dim sTransportista As String
Dim nNeto As Double
Dim nTara As Double
Dim nTotal As Double
Dim sitem As String
Dim sproducto As String

Dim Numguia As String

Sub LlenaCombos()
    With cboMotivo
         Isql = "SELECT Codigo, Descripcion FROM vMotivoTraslado where lActivo=1"
         Set RsMotivo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotivo
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboVehiculo
         Isql = "SELECT Codigo, Descripcion FROM vVehiculo where lActivo=1"
         Set RsVehiculo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsVehiculo
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboChofer
         Isql = "SELECT Codigo, Descripcion FROM vChofer where lActivo=1"
         Set RsChofer = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsChofer
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboEstado
         Isql = "SELECT Codigo, Descripcion FROM vEstadoDocumento"
         Set RsEstado = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstado
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
End Sub

Private Sub cboChofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub cboMotivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub cboVehiculo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub cmdAgproducto_Click()
    cmdBusca(0).Enabled = False
    txtDocumento.Enabled = False
    Isql = "select codigo,tipoproducto,descripcion,grupo,subgrupo from vproducto where lactivo=1"
    frmBusca.cboCriterio.Enabled = True
    frmBusca.nPredeterm = 0
    
    Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "codigo", 800, 0, 0, "", _
                                          "Tipo Producto", 2, "tipoproducto", 1200, 0, 0, "", _
                                          "Descripción", 2, "descripcion", 3200, 0, 0, "", _
                                          "Grupo", 2, "grupo", 1100, 0, 0, "", _
                                          "Subgrupo", 2, "subgrupo", 1000, 0, 0, "")
                                          
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    
    sproducto = sCodigo
    If sTemporal <> "" Then
       sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from " & sTemporal & " ", Cn), 3)
    End If
    'inserta en los productos q no tienen guias asignadas
    Isql = "insert into " & sTemporal & " (tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad, Estado) " & _
           "select " & sitem & " as titem,tcodigoproducto,nprecioventa,'1' as nCantidad,nPrecioventa,'Kgr.' as Unidad,'' as estado from tproducto where tcodigoproducto='" & Trim(sproducto) & "'"
    Cn.Execute Isql
    RsDetalle1.Requery
    RsDetalle2.Requery
End Sub

Private Sub cmdBusca_Click(Index As Integer)
    sDescrip = ""
    If Index = 0 Then
        cmdAgproducto.Enabled = False
        cmdEliproducto.Enabled = False
        cmdBusca(2).Enabled = False
        
        Isql = "SELECT distinct dbo.MDOCUMENTO.tDocumento as Descripcion, dbo.MDOCUMENTO.fFecha, dbo.vCliente.Descripcion as Cliente, dbo.vCliente.tIdentidad, dbo.MDOCUMENTO.nVenta, vCliente.Codigo AS Codigo " & _
               "FROM   dbo.MDOCUMENTO INNER JOIN dbo.DDOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento LEFT OUTER JOIN dbo.vCliente ON dbo.MDOCUMENTO.tCodigoCliente = dbo.vCliente.Codigo " & _
               "WHERE (ISNULL(tCodigoCliente,'')<>'' or len(ltrim(tCodigoCliente)) > 0) and isnull(tguia,'')='' and tEstadoDocumento<>'04'"
               
        frmBusca.cboCriterio.Enabled = True
        frmBusca.nPredeterm = 0
        Call ConfGrilla(5, frmBusca.grdGrilla, "Documento", 2, "Descripcion", 1500, 0, 0, "", _
                                               "Fecha", 2, "fFecha", 1100, 0, 0, "dd/MM/yyyy", _
                                               "Monto", 2, "nVenta", 1000, 1, 0, "###,##0.00", _
                                               "Identificador", 2, "tIdentidad", 1100, 0, 0, "", _
                                               "Razón Comercial", 2, "Cliente", 2600, 0, 0, "")
        frmBusca.Show vbModal
        If Not wEnter Then
           Exit Sub
        End If
        
        Isql = "delete from " & sTemporal
        Cn.Execute Isql
        
        sCliente = sCodigo
        txtDocumento.Text = sDescrip
        txtDestinatario.Text = Calcular("select Descripcion as codigo from vCliente where Codigo='" & sCliente & "'", Cn)
        
        'inserta en los productos q no tienen guias asignadas
        Isql = "insert into " & sTemporal & " (tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad, Estado) " & _
               "select tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, 'Kgr.', ' ' from DDOCUMENTO where tDocumento='" & txtDocumento.Text & "'AND (ISNULL(tGuia,'')='' or len(ltrim(tGuia)) = 0)"
        Cn.Execute Isql
        RsDetalle1.Requery
        RsDetalle2.Requery
   ElseIf Index = 1 Then
        sVar1 = "Transportista"
        Isql = "Select Codigo, Descripcion, tIdentidad, tDireccion from vCLIENTE order by Descripcion"
        frmBusca.cboCriterio.Enabled = True
        frmBusca.nPredeterm = 2
        Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 0, 0, "", _
                                               "Identificador", 2, "tIdentidad", 1100, 0, 0, "", _
                                               "Razon Social", 2, "Descripcion", 2500, 0, 0, "", _
                                               "Dirección", 2, "tDireccion", 2500, 0, 0, "")
        frmBusca.Show vbModal
        If Not wEnter Then
           Exit Sub
        End If
        sTransportista = sCodigo
        txtTransportista.Text = sDescrip
        txtDomicilioDestino.Text = Calcular("select tDireccion as Codigo from vCliente where Codigo='" & sTransportista & "'", Cn)
    Else
        sVar1 = "Transportista"
        Isql = "Select Codigo, Descripcion, tIdentidad, tDireccion from vCLIENTE order by Descripcion"
        frmBusca.cboCriterio.Enabled = True
        frmBusca.nPredeterm = 2
        Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1000, 0, 0, "", _
                                               "Identificador", 2, "tIdentidad", 1100, 0, 0, "", _
                                               "Razon Social", 2, "Descripcion", 2500, 0, 0, "", _
                                               "Dirección", 2, "tDireccion", 2500, 0, 0, "")
        frmBusca.Show vbModal
        If Not wEnter Then
           Exit Sub
        End If
        txtDestinatario.Text = sDescrip
        sCliente = sCodigo
    End If
End Sub

Private Sub cmdEliproducto_Click()
    If RsDetalle1.RecordCount > 0 Then
        Isql = "delete from " & sTemporal & " where tcodigoproducto='" & Trim(RsDetalle1!tCodigoProducto) & "'"
        Cn.Execute Isql
        RsDetalle1.Requery
        RsDetalle2.Requery
    End If
    If grdDetalle1.Row < 0 Then
        cmdBusca(0).Enabled = True
        txtDocumento.Enabled = True
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmGuias.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmGuias.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmGuias.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmGuias.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmGuias.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmGuias.grdGrilla
    End Select
    Asignar
    CalcularMontos
    cmdTexto.Caption = "Registro " & IIf(frmGuias.RsCabecera.RecordCount = 0, 0, frmGuias.RsCabecera.AbsolutePosition) & " de " & frmGuias.RsCabecera.RecordCount
    
End Sub

Private Sub cmdMovimiento_Click(Index As Integer)
    Screen.MousePointer = vbHourglass

    Select Case Index
           Case Is = 0  ' Pasar Todas
                Cn.Execute "update " & sTemporal & " set Estado ='X'"

           Case Is = 1  ' Quitar Todas
                Cn.Execute "update " & sTemporal & " set Estado =' '"

           Case Is = 2  ' Pasar una
                If RsDetalle1.RecordCount <> 0 Then
                    Isql = "update " & sTemporal & " set Estado ='X' where tItem = '" & RsDetalle1!tItem & "'"
                  Cn.Execute Isql
                End If

           Case Is = 3  ' Quitar una
                If RsDetalle2.RecordCount <> 0 Then
                   Cn.Execute "update " & sTemporal & " set Estado =' ' where tItem = '" & RsDetalle2!tItem & "'"
                End If
    End Select
    RsDetalle1.Requery
    RsDetalle2.Requery
    nNeto = Calcular("select sum(nCantidad) as codigo from " & sTemporal & " where Estado='X'", Cn)

    CalcularMontos
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Salir
               Unload Me
          
          Case Is = 1 'Agregar
               Sw = True
               ActivarBotones True
               Blanquear Me
               Inicializa
               Cn.Execute "delete " & sTemporal
               RsDetalle1.Requery
               RsDetalle2.Requery
               
          Case Is = 2 'Grabar
               Dim nCorrela As String
               Dim nPos As Integer
                                        
               'Chequea Datos
               If txtSerie.Text = "" Then MsgBox "Ingrese el número de serie", vbExclamation, sMensaje: txtSerie.SetFocus: Exit Sub
               If txtCorrelativo.Text = "" Then MsgBox "Ingrese el correlativo", vbExclamation, sMensaje: txtCorrelativo.SetFocus: Exit Sub
               If cboMotivo.BoundText = "" Then MsgBox "Ingrese el motivo de traslado", vbExclamation, sMensaje: cboMotivo.SetFocus: Exit Sub
               If cboVehiculo.BoundText = "" Then MsgBox "Ingrese el vehiculo", vbExclamation, sMensaje: cboVehiculo.SetFocus: Exit Sub
               If txtTransportista.Text = "" Then MsgBox "Ingrese el transportista", vbExclamation, sMensaje: cmdBusca(1).SetFocus: Exit Sub
               If cboChofer.BoundText = "" Then MsgBox "Ingrese el chofer", vbExclamation, sMensaje: cboChofer.SetFocus: Exit Sub
                    
               If Sw Then   'Grabar una nueva guía
                  'Verifica si existe el correlativo
                  Isql = "select tGuia as codigo from MGUIA where tGuia='G" & txtSerie.Text & txtCorrelativo.Text & "'"
                  If Calcular(Isql, Cn) <> "0" Then
                     MsgBox "Número de Guia existente", vbCritical, sMensaje
                     Exit Sub
                  End If
                  
                  Sw = False
                  'erwin
                  cboEstado.BoundText = "01"
                  Isql = "INSERT INTO MGUIA (tGuia, fFecha, tDestinatario, tTransportista, tMotivo, tDocumento, tEstado, tChofer, tVehiculo, tObservacion, tDomicilioPartida, tDomicilioDestino, nPeso, nTara, nTotal, tUsuario, fRegistro) Values " & _
                        "('G" & txtSerie.Text & txtCorrelativo.Text & "' " & _
                        ",'" & Format(dtpFecha.value, "yyyy-MM-dd") & "' " & _
                        ",'" & sCliente & "'" & _
                        ",'" & sTransportista & "'" & _
                        ",'" & cboMotivo.BoundText & "'" & _
                        ",'" & txtDocumento & "'" & _
                        ",'" & cboEstado.BoundText & "'" & _
                        ",'" & cboChofer.BoundText & "'" & _
                        ",'" & cboVehiculo.BoundText & "'" & _
                        ",'" & txtObservacion & "'" & _
                        ",'" & txtDomicilioPartida & "'" & _
                        ",'" & txtDomicilioDestino & "'" & _
                        "," & nNeto & _
                        "," & nTara & _
                        "," & nTotal & _
                        ",'" & sUsuario & "'" & _
                        ",GETDATE())"
                    Cn.Execute Isql
                    
                    Isql = "INSERT INTO DGUIA (tGuia, tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad) " & _
                           "select 'G" & txtSerie & txtCorrelativo & "', tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad from " & sTemporal & " where Estado='X'"
                    Cn.Execute Isql
                    
                    Isql = "update DDOCUMENTO set tGuia='G" & txtSerie & txtCorrelativo & "' where tDocumento='" & txtDocumento & "' AND tItem IN (select tItem from dguia where tGuia='G" & txtSerie & txtCorrelativo & "' )"
                    Cn.Execute Isql
                                        
                    frmGuias.RsCabecera.Sort = "tGuia ASC"
                    frmGuias.RsCabecera.Requery
                    frmGuias.RsCabecera.MoveLast
                    MsgBox "Registro Guardado", vbInformation, sMensaje

                    ActivarBotones True
                    cmdTexto.Caption = "Registro " & IIf(frmGuias.RsCabecera.RecordCount = 0, 0, frmGuias.RsCabecera.AbsolutePosition) & " de " & frmGuias.RsCabecera.RecordCount
                    
               Else 'procedimiento para actualizar los datos de una guía
                    Isql = "UPDATE MGUIA SET " & _
                            "fFecha='" & Format(dtpFecha.value, "yyyy-MM-dd") & "'" & _
                            ", tDestinatario='" & sCliente & "'" & _
                            ", tTransportista='" & sTransportista & "'" & _
                            ", tMotivo='" & cboMotivo.BoundText & "'" & _
                            ", tChofer='" & cboChofer.BoundText & "'" & _
                            ", tVehiculo='" & cboVehiculo.BoundText & "'" & _
                            ", tObservacion='" & txtObservacion & "'" & _
                            ", tDomicilioPartida='" & txtDomicilioDestino & "'" & _
                            ", nPeso=" & nNeto & _
                            ", nTara=" & nTara & _
                            ", nTotal=" & nTotal & _
                            ", tUsuario='" & sUsuario & "'" & _
                            " WHERE tGuia='G" & txtSerie.Text & txtCorrelativo.Text & "'"
                    Cn.Execute Isql

                    Isql = "DELETE FROM DGUIA WHERE tGUIA='G" & txtSerie.Text & txtCorrelativo.Text & "'"
                    Cn.Execute Isql
                    
                    Isql = "INSERT INTO DGUIA (tGuia, tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad) " & _
                           "select 'G" & txtSerie & txtCorrelativo & "', tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad from " & sTemporal & " where Estado='X'"
                    Cn.Execute Isql
                    
                    Isql = "update DDOCUMENTO set tGuia='' where tDocumento='" & txtDocumento.Text & "'"
                    Cn.Execute Isql
                    
                    Isql = "update DDOCUMENTO set tGuia='G" & txtSerie & txtCorrelativo & "' where tDocumento='" & txtDocumento.Text & "' and tDocumento+tItem in (select tDocumento+tItem from " & sTemporal & " where Estado='X')"
                    Cn.Execute Isql
                                                           
                    nPos = frmGuias.RsCabecera.AbsolutePosition
                    frmGuias.RsCabecera.Requery
                    frmGuias.RsCabecera.AbsolutePosition = nPos
                    MsgBox "Registro Modificado", vbInformation, sMensaje
                    ActivarBotones True
               End If
          
            Case Is = 3 'cancelar
                If frmGuias.RsCabecera.RecordCount = 0 Then
                    Exit Sub
                End If
                Asignar
                CalcularMontos
                ActivarBotones True
                
            Case Is = 4 'Anular
                If MsgBox("Seguro de Anular", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje) = vbYes Then
                    Isql = "update Mguia set testado='04' where tguia='" & Trim(frmGuias.RsCabecera!tGuia) & "'"
                    Cn.Execute Isql
                    Isql = "Update ddocumento set tguia='' where tdocumento='" & Trim(txtDocumento) & "'"
                    Cn.Execute Isql
                    frmGuias.RsCabecera.Requery
                    Asignar
                End If
                
            Case Is = 5 'Eliminar
                 If frmGuias.RsCabecera.RecordCount = 0 Then
                    Exit Sub
                 End If
               
                 If frmGuias.RsCabecera!tEstado <> "01" Then
                    MsgBox "La guia no esta Emitida, no se puede Eliminar", vbCritical, sMensaje
                    Exit Sub
                 End If
                 Numguia = "G" & txtSerie & txtCorrelativo
                 If MsgBox("Seguro de Eliminar la Guia " & Numguia & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                    Exit Sub
                 End If
                
                 Isql = "DELETE FROM MGUIA WHERE tGuia='" & Numguia & "'"
                 Cn.Execute Isql
                 Isql = "DELETE FROM DGUIA WHERE tGuia='" & Numguia & "'"
                 Cn.Execute Isql
                 Isql = "UPDATE  DDOCUMENTO SET tGuia=NULL WHERE tDocumento='" & txtDocumento & "' AND tGuia='" & Numguia & "'"
                 Cn.Execute Isql
                 frmGuias.RsCabecera.Requery
                 If frmGuias.RsCabecera.RecordCount > 0 Then
                    frmGuias.RsCabecera.MoveLast
                    Asignar
                    'cmdTexto.Caption = "Registro " & IIf(frmGuias.RsCabecera.RecordCount = 0, 0, frmGuias.RsCabecera.AbsolutePosition) & " de " & frmGuias.RsCabecera.RecordCount
                 Else
                    cmdOpcion_Click 1
                    CalcularMontos
                    
                    Sw = True
                 End If
            
            Case Is = 6 ' Imprime Documento
                 Isql = "SELECT dbo.MGUIA.tGuia, dbo.MGUIA.tDocumento , dbo.MGUIA.fFecha, dbo.vCliente.Descripcion AS Destinatario, dbo.vCliente.tIdentidad AS RucDestinatario, dbo.vCliente.tDireccion AS DireccionDestinatario, vCliente_1.Descripcion AS Transportista, vCliente_1.tIdentidad AS RucTransportista, " & _
                        "vCliente_1.tDireccion AS DireccionTransportista, dbo.vChofer.Descripcion AS Chofer, dbo.vChofer.tResumido AS Licencia, dbo.vVehiculo.Descripcion AS Vehiculo, dbo.vVehiculo.Marca, dbo.vVehiculo.Serie, dbo.MGUIA.tMotivo, dbo.MGUIA.tObservacion, dbo.MGUIA.fRegistro, " & _
                        "dbo.MGUIA.tDomicilioPartida, dbo.MGUIA.nPeso, dbo.MGUIA.nTara, dbo.MGUIA.nTotal, dbo.vProducto.Descripcion AS Producto, dbo.vProducto.tResumido AS CodigoProducto, dbo.DGUIA.nCantidad, dbo.DGUIA.tUnidad " & _
                        "FROM dbo.vProducto RIGHT OUTER JOIN dbo.DGUIA ON dbo.vProducto.Codigo = dbo.DGUIA.tCodigoProducto RIGHT OUTER JOIN dbo.MGUIA LEFT OUTER JOIN dbo.vCliente vCliente_1 ON dbo.MGUIA.tTransportista = vCliente_1.Codigo LEFT OUTER JOIN " & _
                        "dbo.vCliente ON dbo.MGUIA.tDestinatario = dbo.vCliente.Codigo LEFT OUTER JOIN dbo.vVehiculo ON dbo.MGUIA.tVehiculo = dbo.vVehiculo.Codigo LEFT OUTER JOIN dbo.vChofer ON dbo.MGUIA.tChofer = dbo.vChofer.Codigo ON dbo.DGUIA.tGuia = dbo.MGUIA.tGuia " & _
                        "WHERE dbo.MGUIA.tGuia='G" & txtSerie.Text & txtCorrelativo.Text & "'"
                 Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                 If RsImpresion.RecordCount = 0 Then
                    LimpiaRs
                    MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                 Else
                   Screen.MousePointer = vbHourglass
                   'Configura la Impresora
                   Dim sImpresora As String
                   sImpresora = Calcular("select tImpresora as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and Cliente =1", Cn)
                   If sImpresora = "0" Then
                      Imprimir ("000")
                   Else
                      Imprimir (sImpresora)
                   End If
                   Printer.FontName = sFont
                   Printer.FontBold = False
                   imprimeGuia RsImpresion, sEmpresa
                   Cn.Execute "UPDATE MGUIA SET tEstado='05' WHERE tGuia='G" & txtSerie.Text & txtCorrelativo.Text & "'"
                   Screen.MousePointer = vbDefault
                End If
                
            Case Is = 7
            Case Is = 8
   End Select
End Sub

Public Sub ActivarBotones(Activa As Boolean)
    cmdMovimiento(0).Enabled = Activa
    cmdMovimiento(1).Enabled = Activa
    cmdMovimiento(2).Enabled = Activa
    cmdMovimiento(3).Enabled = Activa
            
    'cmdOpcion(1).Enabled = Activa      ' Agregar Nuevo
    cmdOpcion(2).Enabled = Activa       ' Grabar Documento
    cmdOpcion(3).Enabled = Activa       ' Cancelar
    cmdOpcion(4).Enabled = Not Activa   ' Anular
    cmdOpcion(5).Enabled = Activa       ' Eliminar
    'cmdOpcion(6).Enabled = Activa      ' Emitir
    
    cmdAgproducto.Enabled = Activa
    cmdEliproducto.Enabled = Activa
    
    txtSerie.Enabled = Not Activa
    txtCorrelativo.Enabled = Not Activa
    dtpFecha.Enabled = Activa
    cboMotivo.Enabled = Activa
    cboChofer.Enabled = Activa
    cboVehiculo.Enabled = Activa
    txtObservacion.Enabled = Activa
    txtDomicilioPartida.Enabled = Activa
    txtNeto.Enabled = Activa
    txtTara.Enabled = Activa
    txtTotal.Enabled = Activa
   
    cmdBusca(0).Enabled = Not Activa
    cmdBusca(1).Enabled = Activa
    grdDetalle2.Enabled = Activa
    grdDetalle1.Enabled = Activa
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
       
    sTemporal = dbTemporal(sCaja, 8, "tGuia", "nVarChar(15)", _
                                     "tItem", "nVarChar(3)", _
                                     "tCodigoProducto", "nVarChar(7)", _
                                     "nPrecioVenta", "Float", _
                                     "nCantidad", "Float", _
                                     "nVenta", "Float", _
                                     "tUnidad", "nVarChar(20)", _
                                     "Estado", "nVarChar(1)")
                                    
    Isql = "SELECT [" & sTemporal & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
           "FROM TPRODUCTO INNER JOIN [" & sTemporal & "] ON TPRODUCTO.tCodigoProducto = [" & sTemporal & "].tCodigoProducto where Estado=' '"
           
    Set RsDetalle1 = Lib.OpenRecordset(Isql, Cn)
    Call ConfGrilla(5, grdDetalle1, "Producto", 2, "Producto", 1600, 0, 0, "", _
                                    "Precio", 2, "nPrecioVenta", 810, 1, 0, "###,###,##0.00", _
                                    "Cant", 2, "nCantidad", 640, 1, 0, "#,##0.00", _
                                    "Unid", 2, "tUnidad", 600, 0, 0, "", _
                                    "SubTotal", 2, "nVenta", 830, 1, 0, "###,###,##0.00")
    Set grdDetalle1.DataSource = RsDetalle1
   
    Isql = "SELECT [" & sTemporal & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
           "FROM TPRODUCTO INNER JOIN [" & sTemporal & "] ON TPRODUCTO.tCodigoProducto = [" & sTemporal & "].tCodigoProducto where Estado='X'"
    Set RsDetalle2 = Lib.OpenRecordset(Isql, Cn)
    Call ConfGrilla(5, grdDetalle2, "Producto", 2, "Producto", 1600, 0, 0, "", _
                                    "Precio", 2, "nPrecioVenta", 810, 1, 0, "###,###,##0.00", _
                                    "Cant", 2, "nCantidad", 640, 1, 0, "#,##0.00", _
                                    "Unid", 2, "tUnidad", 600, 0, 0, "", _
                                    "SubTotal", 2, "nVenta", 830, 1, 0, "###,###,##0.00")
    
    Set grdDetalle2.DataSource = RsDetalle2
      
    LlenaCombos
                            
    If Sw Then
        Blanquear Me
        ActivarBotones True
        Inicializa
    Else
        Asignar
        CalcularMontos
    End If
    
   Screen.MousePointer = vbDefault
   cmdTexto.Caption = "Registro " & IIf(frmGuias.RsCabecera.RecordCount = 0, 0, frmGuias.RsCabecera.AbsolutePosition) & " de " & frmGuias.RsCabecera.RecordCount

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmGuiasDetalle = Nothing
   Cn.Execute "drop table " & sTemporal
End Sub

Private Sub grdDetalle1_DblClick()
    If RsDetalle1.RecordCount > 0 Then
       Dim und As String
       Dim pos As Integer
       und = InputBox("Ingrese la nueva cantidad del producto", "Nueva cantidad")
       If Not IsNumeric(und) Then
          MsgBox "Se debe ingresar un dato numérico", vbExclamation, sMensaje
          Exit Sub
       End If
       If und <> "" Then
           pos = RsDetalle1.AbsolutePosition
           Isql = "UPDATE " & sTemporal & " SET nCantidad='" & und & "' WHERE tItem='" & RsDetalle1!tItem & "' AND Estado=' '"
           Cn.Execute Isql
           RsDetalle1.Requery
           RsDetalle1.AbsolutePosition = pos
       End If
    End If
End Sub

Private Sub grdDetalle2_DblClick()
    If RsDetalle2.RecordCount > 0 Then
       Dim und As String
       Dim pos As Integer
       und = InputBox("Ingrese la nueva undidad del producto", "Nueva Unidad")
       If und <> "" Then
          pos = RsDetalle2.AbsolutePosition
          Isql = "UPDATE " & sTemporal & " SET tUnidad='" & Mid(und, 1, 4) & "' WHERE tItem='" & RsDetalle2!tItem & "' AND Estado='X'"
          Cn.Execute Isql
          RsDetalle2.Requery
          RsDetalle2.AbsolutePosition = pos
       End If
    End If
End Sub

Private Sub txtTotal_GotFocus()
   txtTotal.SelLength = Len(Trim(txtTotal.Text))
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
   Numerico KeyAscii, txtTotal
End Sub

Private Sub txtTotal_LostFocus()
   nTotal = Val(txtTotal.Text)
   CalcularMontos
End Sub

Private Sub txtCorrelativo_GotFocus()
   txtCorrelativo.SelLength = Len(Trim(txtCorrelativo.Text))
End Sub

Private Sub txtCorrelativo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtCorrelativo_LostFocus()
   txtCorrelativo = Mid("000000000", 1, 9 - Len(Trim(str(Val(txtCorrelativo))))) + Trim(str(Val(txtCorrelativo)))
End Sub

Private Sub txtNeto_GotFocus()
   txtNeto.SelLength = Len(Trim(txtNeto.Text))
End Sub

Private Sub txtNeto_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
   Numerico KeyAscii, txtNeto
End Sub

Private Sub txtNeto_LostFocus()
   nNeto = Val(txtNeto.Text)
   CalcularMontos
End Sub

Private Sub txtSerie_GotFocus()
   txtSerie.SelLength = Len(Trim(txtSerie.Text))
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtSerie_LostFocus()
   txtSerie = Mid("00000", 1, 5 - Len(Trim(str(Val(txtSerie))))) + Trim(str(Val(txtSerie)))
End Sub

Public Sub Inicializa()
    sCliente = ""
    sTransportista = ""
    sUsuario = ""
    txtDomicilioPartida = sDireccion
    nNeto = 0
    nTara = 0
    nTotal = 0
    txtNeto.Text = "0.00"
    txtTara.Text = "0.00"
    txtTotal.Text = "0.00"
    dtpFecha.value = FechaServidor()
    txtSerie.Enabled = True
    txtCorrelativo.Enabled = True
    cmdBusca(0).Enabled = True
    
    cmdOpcion(1).Enabled = False   ' Agregar Nuevo
    cmdOpcion(5).Enabled = False   ' Eliminar
    cmdOpcion(6).Enabled = False   ' Emitir
    
End Sub

Sub Asignar()
   Screen.MousePointer = vbHourglass
  
   With frmGuias.RsCabecera
        If .RecordCount <= 0 Then
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
            txtSerie = IIf(IsNull(Mid(!tGuia, 2, 5)), "", Mid(!tGuia, 2, 5))
            txtCorrelativo = IIf(IsNull(Mid(!tGuia, 7, 9)), "", Mid(!tGuia, 7, 9))
            dtpFecha.value = Format(!fFecha, "dd MM yyyy")
            txtDestinatario = IIf(IsNull(!Destinatario), "", !Destinatario)
            txtTransportista = IIf(IsNull(!Transportista), "", !Transportista)
            cboMotivo.BoundText = IIf(IsNull(!tMotivo), "", !tMotivo)
            cboEstado.BoundText = IIf(IsNull(!tEstado), "", !tEstado)
            txtDocumento = IIf(IsNull(!tDocumento), "", !tDocumento)
            cboChofer.BoundText = IIf(IsNull(!tChofer), "", !tChofer)
            cboVehiculo.BoundText = IIf(IsNull(!tVehiculo), "", !tVehiculo)
            txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
            txtDomicilioDestino = IIf(IsNull(!tDomicilioDestino), "", !tDomicilioDestino)
            txtDomicilioPartida = IIf(IsNull(!tDomicilioPartida), "", !tDomicilioPartida)
            txtNeto = Format(IIf(IsNull(!nPeso), "", !nPeso), "###,##0.00")
            txtTara = Format(IIf(IsNull(!nTara), "", !nTara), "###,##0.00")
            txtTotal = Format(IIf(IsNull(!nTotal), "", !nTotal), "###,##0.00")
            sUsuario = IIf(IsNull(!tUsuario), "", !tUsuario)
            sCliente = IIf(IsNull(!tDestinatario), "", !tDestinatario)
            sTransportista = IIf(IsNull(!tTransportista), "", !tTransportista)
            sDocumento = IIf(IsNull(!tDocumento), "", !tDocumento)
            nNeto = txtNeto
            nTara = txtTara
            nTotal = txtTotal
            
            'mostramos los productos q estan en la guía
            Cn.Execute "delete from " & sTemporal
            Isql = "INSERT into " & sTemporal & " (tGuia,tItem,tCodigoProducto,nPrecioVenta,nCantidad,nVenta,tUnidad,Estado) " & _
                   "select dGuia.tGuia, dGuia.tItem, dGuia.tCodigoProducto, dGuia.nPrecioVenta, dGuia.nCantidad, dGuia.nVenta, dGuia.tUnidad, 'X' as Estado " & _
                   "from dGuia where tGuia='" & !tGuia & "'"
            Cn.Execute Isql
            
            'mostramos los productos q no tienen guías solo de las guias q no tienn estado ANULADO
            If Calcular("select tEstado as codigo from mGuia where tGUIA='G" & txtSerie.Text & txtCorrelativo.Text & "'", Cn) = "01" Then
                Isql = "insert into " & sTemporal & " (tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta, tUnidad, Estado) " & _
                        "select tItem, DDOCUMENTO.tCodigoProducto, DDOCUMENTO.nPrecioVenta, nCantidad, nVenta, 'Kgr.' as tUnidad, ' ' as EStado " & _
                        "from DDOCUMENTO LEFT JOIN vProducto ON vProducto.codigo = DDOCUMENTO.tCodigoProducto where tDocumento='" & txtDocumento & "' AND ISNULL(tGuia,'')='' "
                Cn.Execute Isql
            End If
            RsDetalle1.Requery
            RsDetalle2.Requery
            
            Sw = False
            If cboEstado.BoundText = "01" Then
                ActivarBotones True
            Else
                ActivarBotones False
                cmdBusca(0).Enabled = False
            End If
    
  End With
  Screen.MousePointer = vbDefault
End Sub

Public Sub CalcularMontos()
    nTotal = nTara + nNeto
    txtNeto.Text = Format(nNeto, "###,##0.00")
    txtTara.Text = Format(nTara, "###,##0.00")
    txtTotal.Text = Format(nTotal, "###,##0.00")
    
    txtTotal2 = Format(Calcular("select sum(nVenta) AS codigo from " & sTemporal & " where Estado='X'", Cn), "###,##0.00")
    txtTotal1 = Format(Calcular("select sum(nVenta) AS codigo from " & sTemporal & " where Estado=' '", Cn), "###,##0.00")
       
End Sub

Private Sub txtTara_GotFocus()
   txtTara.SelLength = Len(Trim(txtTara.Text))
End Sub

Private Sub txtTara_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
   Numerico KeyAscii, txtTara
End Sub

Private Sub txtTara_LostFocus()
    nTara = Val(txtTara.Text)
    CalcularMontos
End Sub

