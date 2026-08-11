VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTipoPedidoDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   -570
   ClientTop       =   1875
   ClientWidth     =   9375
   FillColor       =   &H8000000F&
   Icon            =   "frmTipoPedidoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9375
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9315
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4290
      Width           =   9375
      Begin VB.CommandButton cmdSalir 
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
         Left            =   7920
         Picture         =   "frmTipoPedidoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   15
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   6705
         Picture         =   "frmTipoPedidoDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   15
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   645
         Left            =   30
         ScaleHeight     =   585
         ScaleWidth      =   6000
         TabIndex        =   28
         Top             =   60
         Width           =   6060
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoPedidoDetalle.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoPedidoDetalle.frx":0FA8
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoPedidoDetalle.frx":14EA
            Style           =   1  'Graphical
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5520
            Picture         =   "frmTipoPedidoDetalle.frx":1A2C
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5040
            Picture         =   "frmTipoPedidoDetalle.frx":1F6E
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4560
            Picture         =   "frmTipoPedidoDetalle.frx":24B0
            Style           =   1  'Graphical
            TabIndex        =   3
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
            Left            =   1485
            TabIndex        =   29
            Top             =   150
            Width           =   3000
         End
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
      Height          =   4260
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9300
      Begin TabDlg.SSTab SSTab1 
         Height          =   3165
         Left            =   225
         TabIndex        =   30
         Top             =   975
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   5583
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos Principales"
         TabPicture(0)   =   "frmTipoPedidoDetalle.frx":29F2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Otros Datos"
         TabPicture(1)   =   "frmTipoPedidoDetalle.frx":2A0E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(2)=   "Frame3"
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame5 
            Caption         =   "Mantenimiento de Origen de Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   -70800
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   4455
            Begin VB.CommandButton cmdModOV 
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
               Index           =   1
               Left            =   3720
               Picture         =   "frmTipoPedidoDetalle.frx":2A2A
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Modifica"
               Top             =   2160
               Width           =   555
            End
            Begin VB.CommandButton cmdNewOV 
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
               Left            =   3120
               Picture         =   "frmTipoPedidoDetalle.frx":2B2C
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Agrega"
               Top             =   2160
               Width           =   555
            End
            Begin TrueOleDBGrid80.TDBGrid gdrOrigenVenta 
               Height          =   1875
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   4245
               _ExtentX        =   7488
               _ExtentY        =   3307
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
               Splits(0).ScrollBars=   2
               Splits(0).AllowColMove=   -1  'True
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).FetchRowStyle=   -1  'True
               Splits(0).DividerStyle=   3
               Splits(0).DividerColor=   32768
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
               TabAction       =   3
               MultipleLines   =   0
               CellTipsWidth   =   0
               InsertMode      =   0   'False
               MultiSelect     =   0
               GroupByCaption  =   ""
               DeadAreaBackColor=   -2147483637
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
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=13,.bold=-1,.fontsize=825,.italic=0"
               _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(10)  =   ":id=2,.fontname=Arial"
               _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=3,.fontname=Calibri"
               _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
               _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2"
               _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(22)  =   "Splits(0).Style:id=30,.parent=1,.namedParent=33"
               _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=49,.parent=4"
               _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=31,.parent=2"
               _StyleDefs(25)  =   "Splits(0).FooterStyle:id=32,.parent=3"
               _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=43,.parent=5"
               _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=45,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
               _StyleDefs(28)  =   "Splits(0).EditorStyle:id=44,.parent=7"
               _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=46,.parent=8"
               _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=47,.parent=9"
               _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=48,.parent=10"
               _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=50,.parent=11"
               _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=51,.parent=12"
               _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=55,.parent=30"
               _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=52,.parent=31,.alignment=0"
               _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=53,.parent=32"
               _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=54,.parent=44"
               _StyleDefs(38)  =   "Named:id=33:Normal"
               _StyleDefs(39)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
               _StyleDefs(40)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(41)  =   ":id=33,.fontname=Calibri"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
               _StyleDefs(44)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
               _StyleDefs(45)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(46)  =   ":id=34,.fontname=Arial Black"
               _StyleDefs(47)  =   "Named:id=35:Footing"
               _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(49)  =   "Named:id=36:Selected"
               _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
               _StyleDefs(51)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(52)  =   ":id=36,.fontname=Calibri"
               _StyleDefs(53)  =   "Named:id=37:Caption"
               _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
               _StyleDefs(55)  =   "Named:id=38:HighlightRow"
               _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
               _StyleDefs(57)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(58)  =   ":id=38,.fontname=Calibri"
               _StyleDefs(59)  =   "Named:id=39:EvenRow"
               _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
               _StyleDefs(61)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(62)  =   ":id=39,.fontname=Calibri"
               _StyleDefs(63)  =   "Named:id=40:OddRow"
               _StyleDefs(64)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
               _StyleDefs(65)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(66)  =   ":id=40,.fontname=Calibri"
               _StyleDefs(67)  =   "Named:id=41:RecordSelector"
               _StyleDefs(68)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
               _StyleDefs(69)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(70)  =   ":id=41,.fontname=Calibri"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
               _StyleDefs(73)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(74)  =   ":id=42,.fontname=Calibri"
               _StyleDefs(75)  =   "Named:id=13:CabeceraTitulo"
               _StyleDefs(76)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
               _StyleDefs(77)  =   ":id=13,.strikethrough=0,.charset=0"
               _StyleDefs(78)  =   ":id=13,.fontname=Arial Black"
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   -70800
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   4455
            Begin VB.CheckBox chkVisualizaDespachador 
               Caption         =   "Ver en despachador"
               Height          =   255
               Left            =   360
               TabIndex        =   60
               Top             =   2040
               Width           =   1890
            End
            Begin VB.TextBox txtCodAPI 
               Height          =   285
               Left            =   3360
               TabIndex        =   52
               Top             =   1560
               Width           =   975
            End
            Begin VB.CheckBox chkDeliveryOV 
               Caption         =   "Delivery"
               Height          =   225
               Left            =   360
               TabIndex        =   51
               Top             =   1560
               Width           =   1170
            End
            Begin VB.TextBox txtbotonOV 
               Height          =   285
               Left            =   3360
               TabIndex        =   49
               Top             =   1200
               Width           =   975
            End
            Begin VB.CheckBox chkVisualizaOV 
               Caption         =   "Visualiza"
               Height          =   225
               Left            =   360
               TabIndex        =   47
               Top             =   1320
               Width           =   1170
            End
            Begin VB.CheckBox ChkActivoOV 
               Caption         =   "Activo"
               Height          =   225
               Left            =   360
               TabIndex        =   43
               Top             =   1800
               Width           =   1170
            End
            Begin VB.CommandButton cmdSalirOV 
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
               Height          =   495
               Left            =   3480
               Picture         =   "frmTipoPedidoDetalle.frx":305E
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   2040
               Width           =   810
            End
            Begin VB.CommandButton GrabarOV 
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
               Height          =   495
               Left            =   2640
               Picture         =   "frmTipoPedidoDetalle.frx":3150
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   2040
               Width           =   810
            End
            Begin VB.TextBox txtCodOV 
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
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   360
               Width           =   1110
            End
            Begin VB.TextBox txtDescripcionOV 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   37
               Top             =   720
               Width           =   3255
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Cod ApiWeb:"
               Height          =   195
               Index           =   7
               Left            =   2280
               TabIndex        =   53
               Top             =   1560
               Width           =   945
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Boton :"
               Height          =   195
               Index           =   6
               Left            =   2760
               TabIndex        =   50
               Top             =   1200
               Width           =   510
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Código :"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   40
               Top             =   405
               Width           =   585
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Descripción :"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   38
               Top             =   795
               Width           =   930
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   -74850
            TabIndex        =   33
            Top             =   375
            Width           =   4005
            Begin VB.TextBox txtEnlaceSunat 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
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
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1005
               Width           =   1575
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
               Height          =   330
               Left            =   3255
               Picture         =   "frmTipoPedidoDetalle.frx":3682
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   960
               Width           =   630
            End
            Begin VB.TextBox txtEnlaceContable2 
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
               Left            =   1650
               MaxLength       =   20
               TabIndex        =   22
               Top             =   600
               Width           =   2250
            End
            Begin VB.TextBox txtEnlaceContable1 
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
               Left            =   1650
               MaxLength       =   20
               TabIndex        =   21
               Text            =   " "
               Top             =   225
               Width           =   2250
            End
            Begin VB.Label lblDescripcion 
               ForeColor       =   &H00FF0000&
               Height          =   975
               Left            =   1560
               TabIndex        =   57
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Enlace Sunat :"
               Height          =   195
               Index           =   8
               Left            =   480
               TabIndex        =   56
               Top             =   1005
               Width           =   1050
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Enlace Contable 2 :"
               Height          =   195
               Index           =   3
               Left            =   150
               TabIndex        =   35
               Top             =   675
               Width           =   1395
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Enlace Contable 1 :"
               Height          =   195
               Index           =   2
               Left            =   150
               TabIndex        =   34
               Top             =   300
               Width           =   1395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Activaciones "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            Left            =   150
            TabIndex        =   32
            Top             =   375
            Width           =   4005
            Begin VB.CheckBox chklCuentaXCobrar 
               Caption         =   "Activar Cuentas por Cobrar"
               Height          =   375
               Left            =   120
               TabIndex        =   61
               Top             =   2040
               Width           =   2655
            End
            Begin VB.CheckBox chklMobileInforest 
               Caption         =   "Activar Mobile Inforest"
               Height          =   375
               Left            =   120
               TabIndex        =   59
               Top             =   1440
               Width           =   2055
            End
            Begin VB.CheckBox chkActivo 
               Caption         =   "Activo"
               Height          =   225
               Left            =   120
               TabIndex        =   13
               Top             =   1800
               Width           =   2490
            End
            Begin VB.CheckBox chkCentralPedido 
               Caption         =   "Canal para Central de Pedidos "
               Height          =   225
               Left            =   120
               TabIndex        =   11
               Top             =   930
               Width           =   2775
            End
            Begin VB.CheckBox chkDelivery 
               Caption         =   "Canal para Delivery"
               Height          =   225
               Left            =   120
               TabIndex        =   12
               Top             =   1245
               Width           =   2910
            End
            Begin VB.CheckBox chkActivaMotorizado 
               Caption         =   "Activar Motorizado"
               Height          =   225
               Left            =   120
               TabIndex        =   10
               Top             =   585
               Width           =   2730
            End
            Begin VB.CheckBox chkActivaMozo 
               Caption         =   "Activar Mozo"
               Height          =   225
               Left            =   120
               TabIndex        =   9
               Top             =   255
               Width           =   1890
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Obligatoriedades "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            Left            =   4200
            TabIndex        =   31
            Top             =   375
            Width           =   4485
            Begin VB.CheckBox ChklObligaOrigenVenta 
               Caption         =   "Obligatoriedad de Origen de Venta"
               Height          =   255
               Left            =   360
               TabIndex        =   58
               Top             =   2300
               Width           =   3015
            End
            Begin VB.CheckBox chkObligaClienteFrecuente 
               Caption         =   "Obligatoriedad Cliente Frecuente"
               Height          =   225
               Left            =   360
               TabIndex        =   19
               Top             =   1720
               Width           =   3210
            End
            Begin VB.CheckBox chkObligaMesa 
               Caption         =   "Obligatoriedad de Mesa"
               Height          =   225
               Left            =   360
               TabIndex        =   16
               Top             =   810
               Width           =   2490
            End
            Begin VB.CheckBox chkObligaMozo 
               Caption         =   "Obligatoriedad de Mozo"
               Height          =   225
               Left            =   360
               TabIndex        =   14
               Top             =   250
               Width           =   2250
            End
            Begin VB.CheckBox chkObligaPax 
               Caption         =   "Obligatoriedad de Pax"
               Height          =   225
               Left            =   360
               TabIndex        =   17
               Top             =   1110
               Width           =   1890
            End
            Begin VB.CheckBox chkObligaMotorizado 
               Caption         =   "Obligatoriedad de Motorizado"
               Height          =   225
               Left            =   360
               TabIndex        =   15
               Top             =   520
               Width           =   2490
            End
            Begin VB.CheckBox chkFechaEntregaDelivery 
               Caption         =   "Obligatoriedad de Ingreso de Fecha Entrega Delivery"
               Height          =   225
               Left            =   360
               TabIndex        =   18
               Top             =   1420
               Width           =   4050
            End
            Begin VB.CheckBox chkEntregarA 
               Caption         =   "Obligatoriedad de ""Entregar A"""
               Height          =   225
               Left            =   360
               TabIndex        =   20
               Top             =   2020
               Width           =   3210
            End
         End
      End
      Begin VB.TextBox txtTipoPedido 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   8
         Top             =   600
         Width           =   6870
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido :"
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   26
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   25
         Top             =   285
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmTipoPedidoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sprefijo As String
Dim i As Integer
'Dim sCodigo As String
Dim RsDetalle As ADODB.Recordset
Dim newOV As Boolean
Dim RsOV As Recordset

Private Sub chkActivaMotorizado_Click()
   If chkActivaMotorizado.Value = 0 Then
      chkObligaMotorizado.Value = 0
   End If
End Sub

Private Sub chkActivaMozo_Click()
   If chkActivaMozo.Value = 0 Then
      chkObligaMozo.Value = 0
   End If
End Sub

Private Sub chkCentralPedido_Click()
   If chkCentralPedido.Value = 0 Then
      chkFechaEntregaDelivery.Value = 0
      chkObligaClienteFrecuente.Value = 0
    Else
        chkFechaEntregaDelivery.Value = 1
        chkObligaClienteFrecuente.Value = 1
   End If
   
End Sub

Private Sub chkDelivery_Click()
   If chkDelivery.Value = 0 Then
      chkObligaClienteFrecuente.Value = 0
   End If
   
    If chkDelivery.Value = 1 Then
        chkActivaMozo.Value = 0
        chkActivaMozo.Enabled = 0
        Else
         chkActivaMozo.Enabled = 1
   End If
   
End Sub

Private Sub cmdBusca_Click()
    Isql = "select Codigo, Descripcion from vProductoSunat order by Codigo"
    Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                           "Producto Sunat", 2, "Descripcion", 6100, 0, 0, "")
    frmBusca.nPredeterm = 1
    frmBusca.Show vbModal
    If wEnter Then
       'sTipoProducto = sCodigo
       Me.lblDescripcion.Caption = sDescrip
       Me.txtEnlaceSunat.Text = sCodigo
    End If
End Sub

 Private Sub cmdGrabar_Click()
        If Me.txtTipoPedido.Text = "" And Me.chkActivo.Value = 1 Then
            MsgBox "Debe indicar una Descripción para el Tipo de Pedido", vbInformation, sMensaje
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Calcular("select count(*) as codigo from vTipoPedido where Codigo not in ('" & txtCodigo & "') and lCanalDelivery = 1", Cn) > 0 And Me.chkDelivery.Value = 1 Then
            MsgBox "La Activacion de Canal Delivery ya esta en uso, imposible grabar!!!", vbExclamation
            Me.chkDelivery.Value = 0
            Exit Sub
        End If

'        If RsDetalle!lCanalDelivery = False And ChklObligaOrigenVenta = 1 Then
'            MsgBox "La Activacion de Obligatoriedad de Origen de Venta, solo es para el Canal Delivery, imposible grabar!!!", vbExclamation
'            Me.ChklObligaOrigenVenta.Value = 0
'            Exit Sub
'        End If

        nPos = frmTipoPedido.RsCabecera.Bookmark
                  
        If chkActivo.Value = 1 Then
           Select Case frmTipoPedido.RsCabecera.Fields(0)
                  Case "01": sBoton1 = txtTipoPedido.Text
                  Case "02": sBoton2 = txtTipoPedido.Text
                  Case "03": sBoton3 = txtTipoPedido.Text
                  Case "04": sBoton4 = txtTipoPedido.Text
                  Case "05": sBoton5 = txtTipoPedido.Text
           End Select
        Else
           Select Case frmTipoPedido.RsCabecera.Fields(0)
                  Case "01": sBoton1 = ""
                  Case "02": sBoton2 = ""
                  Case "03": sBoton3 = ""
                  Case "04": sBoton4 = ""
                  Case "05": sBoton5 = ""
            End Select
        End If
                
        If Calcular("select count(*)  as codigo from TCANALVENTA where tCodigoCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'", Cn) > 0 Then
        
            Isql = " update TCANALVENTA  set " & _
                   "tCodigoSunat = '" & txtEnlaceSunat.Text & "', " & _
                   "tDetallado = '" & txtTipoPedido.Text & "', " & _
                   "tEnlaceContable1 = '" & txtEnlaceContable1.Text & "', " & _
                   "tEnlaceContable2= '" & txtEnlaceContable2.Text & "', " & _
                   " lobligamesa =" & IIf(chkObligaMesa.Value, 1, 0) & ", " & _
                   " lobligapax =" & IIf(chkObligaPax.Value, 1, 0) & ", " & _
                   " lActivamozo =" & IIf(chkActivaMozo.Value, 1, 0) & ", " & _
                   " lobligamozo =" & IIf(chkObligaMozo.Value, 1, 0) & ", " & _
                   " lobligaentregara=" & IIf(Me.chkEntregarA.Value, 1, 0) & ", " & _
                   " lActivamotorizado =" & IIf(chkActivaMotorizado.Value, 1, 0) & ", " & _
                   " lobligamotorizado =" & IIf(chkObligaMotorizado.Value, 1, 0) & ", " & _
                   " lcanalCentralPedidos=" & IIf(chkCentralPedido.Value, 1, 0) & "," & _
                   " lcanaldelivery =" & IIf(chkDelivery.Value, 1, 0) & ", " & _
                   " lActivaCuentasXCobrar =" & IIf(chklCuentaXCobrar.Value, 1, 0) & ", " & _
                   " lObligaIngresoFechaEntrega =" & IIf(chkFechaEntregaDelivery.Value, 1, 0) & ", " & _
                   " lObligaClienteFrecuente =" & IIf(chkObligaClienteFrecuente.Value, 1, 0) & ", " & _
                   " lActivo =" & IIf(chkActivo.Value, 1, 0) & " , " & _
                   " lMobileInforest =" & IIf(chklMobileInforest.Value, 1, 0) & " , " & _
                   " lObligaOrigenVenta =" & IIf(ChklObligaOrigenVenta.Value, 1, 0) & " " & _
                   " where tCodigoCanalVenta ='" & frmTipoPedido.RsCabecera.Fields(0) & "'"
            Cn.Execute Isql
                   
        Else
            Isql = "insert into TCANALVENTA(CodigoCanalVenta, lobligamesa, lobligapax, lactivamozo, lobligamozo, lactivamotorizado,                           lobligamotorizado,                           lcanalCentralPedidos,                       lcanaldelivery,                       lObligaIngresoFechaEntrega,                           lObligaClienteFrecuente, lobligaentregara, tenlacecontable1,tenlacecontable2, tCodigoSunat, lObligaOrigenVenta,lActivaCuentasXCobrar) " & _
            "values('" & frmTipoPedido.RsCabecera.Fields(0) & "',  " & IIf(chkObligaMesa.Value, 1, 0) & ",  " & IIf(chkObligaPax.Value, 1, 0) & ", " & IIf(chkActivaMozo.Value, 1, 0) & "," & IIf(chkObligaMozo.Value, 1, 0) & ", " & IIf(chkActivaMotorizado.Value, 1, 0) & ", " & IIf(chkObligaMotorizado.Value, 1, 0) & ", " & IIf(chkCentralPedido.Value, 1, 0) & ", " & IIf(chkDelivery.Value, 1, 0) & ", " & IIf(chkFechaEntregaDelivery.Value, 1, 0) & ", " & IIf(chkObligaClienteFrecuente.Value, 1, 0) & ", " & IIf(Me.chkEntregarA.Value, 1, 0) & ",'" & Trim(Me.txtEnlaceContable1.Text) & "','" & Trim(Me.txtEnlaceContable2.Text) & "','" & txtEnlaceSunat.Text & "','" & IIf(ChklObligaOrigenVenta.Value, 1, 0) & "','" & IIf(chklCuentaXCobrar.Value, 1, 0) & "')"
            Cn.Execute Isql
        End If

        RsDetalle.Requery
        frmTipoPedido.RsCabecera.Requery
        frmTipoPedido.RsCabecera.Bookmark = nPos
        Screen.MousePointer = vbDefault
        MsgBox "Registro Modificado", vbInformation, sMensaje
End Sub



Private Sub cmdModOV_Click(Index As Integer)
    If RsOV.RecordCount > 0 Then
        Frame4.Visible = True
        Frame5.Visible = False
        newOV = False
        With RsOV
        txtCodOV.Text = IIf(IsNull(!CodOrigenVenta), "", !CodOrigenVenta)
        txtDescripcionOV.Text = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtbotonOV.Text = IIf(IsNull(!Descripcion), "", !Boton)
        chkDeliveryOV.Value = IIf(!Delivery = True, 1, 0)
        chkVisualizaOV.Value = IIf(!Visible = True, 1, 0)
        ChkActivoOV.Value = IIf(!Activo = True, 1, 0)
        ' se agrego esta linea para cambios en canastas
        txtCodAPI.Text = IIf(IsNull(!tCodAPIWEB), "", !tCodAPIWEB)
        Me.chkVisualizaDespachador.Value = IIf(!lVisualizaDespachador = True, 1, 0)
        End With
    End If
            
End Sub
Private Sub cmdNewOV_Click()
    Frame4.Visible = True
    Frame5.Visible = False
    newOV = True

    txtCodOV.Text = ""
    txtDescripcionOV.Text = ""
    txtbotonOV.Text = "0"
    chkVisualizaOV.Value = 0
    chkDeliveryOV.Value = 0
    ChkActivoOV.Value = 1
    ' se agrego
    txtCodAPI.Text = ""
End Sub
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdSalirOV_Click()
    Frame4.Visible = False
    Frame5.Visible = True
    newOV = False
End Sub

Private Sub GrabarOV_Click()
On Error GoTo fin
    Dim CodOV, nCorrela As String
    If txtDescripcionOV.Text = "" Then
        MsgBox "Ingrese un Nombre de Origen de Venta"
        Exit Sub
    End If
    If txtbotonOV.Text = "" Then
        MsgBox "Ingrese un numero entero"
        Exit Sub
    End If
    If newOV = True And txtCodOV.Text = "" Then
        If Calcular("select count(*) as Codigo from vorigenventa where Boton>0 and  Boton=" & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & " and codCanalVenta='" & Me.txtCodigo.Text & "' ", Cn) > 0 Then
            MsgBox "El Boton" & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & " ya esta registrado favor de Corregirlo!!"
            Exit Sub
        End If
        nCorrela = Calcular("select Max(CodOrigenVenta) as Codigo from vorigenventa  ", Cn)
        If IsNull(nCorrela) Or nCorrela = "" Then
            CodOV = "01"
        Else
            CodOV = Lib.Correlativo(nCorrela, 2)
        End If
    ' en el insert se agrego el campo de tCodAPIWEB¨
    ' en la parte de values se agrego el campo de tCodAPIWEB
        Isql = "insert into TORIGENVENTA (tCodigoOrigenVenta,tCodigoCanalVenta,tDescripcion, nBoton, lVisiblePos, lDelivery, lActivo, tCodAPIWEB, lVisualizaDespachador) " & _
                "values ('" & CodOV & "','" & frmTipoPedido.RsCabecera.Fields(0) & "','" & txtDescripcionOV.Text & "'," & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & "," & Me.chkVisualizaOV.Value & "," & Me.chkDeliveryOV.Value & "," & Me.ChkActivoOV.Value & ",'" & txtCodAPI.Text & "'" & "," & Me.chkVisualizaDespachador.Value & ")"
        Cn.Execute Isql
    Else
            If Calcular("select count(*) as Codigo from vorigenventa where Boton>0 and  Boton=" & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & " and CodOrigenVenta<> '" & txtCodOV.Text & "' and codCanalVenta='" & Me.txtCodigo.Text & "'", Cn) > 0 Then
                MsgBox "El Boton" & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & " ya esta registrado favor de Corregirlo!!"
                Exit Sub
            End If
            ' en el update se realizo lo mismo
                Isql = "update TORIGENVENTA set tDescripcion='" & txtDescripcionOV.Text & "', nBoton =" & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & ", lVisiblePos=" & Me.chkVisualizaOV.Value & ", lDelivery=" & Me.chkDeliveryOV.Value & ", lActivo=" & Me.ChkActivoOV.Value & _
                ", tCodAPIWEB='" & Me.txtCodAPI.Text & "', lVisualizaDespachador = " & Me.chkVisualizaDespachador.Value & " " & _
                " where tCodigoOrigenVenta='" & txtCodOV.Text & "'"
               ' "values ('" & CodOV & "','" & frmTipoPedido.RsCabecera.Fields(0) & "','" & txtDescripcionOV.Text & "'," & IIf(txtbotonOV.Text = "", 0, CInt(txtbotonOV.Text)) & "," & Me.chkVisualizaOV.value & "," & Me.chkDeliveryOV.value & "," & Me.ChkActivoOV.value & ")"
        Cn.Execute Isql
    End If
    Frame4.Visible = False
    Frame5.Visible = True
    newOV = False
    RsOV.Requery
    MsgBox "Proceso Correcto"
Exit Sub
fin:
    MsgBox Error, vbCritical, sMensaje
End Sub
Private Sub Form_Load()
    Centrar Me
    Me.Caption = "Mantenimiento de Canales de Venta"
    
    Set RsDetalle = Lib.OpenRecordset("select * from TCANALVENTA where tCodigoCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'", Cn)
    Asignar
    
    Frame5.Visible = True
    
    RsDetalle.MoveFirst
    'If RsDetalle!lCanalDelivery Then
        
        Frame5.Visible = True
        Isql = "select CodOrigenVenta,Descripcion,Boton,Visible,Delivery,Activo, tCodAPIWEB, lVisualizaDespachador  from vorigenventa where CodCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'" '"select * from vOrigenVenta where CodCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'"
        Set RsOV = Lib.OpenRecordset(Isql, Cn)
        Call ConfGrilla(4, gdrOrigenVenta, "Codigo", 2, "CodOrigenVenta", 700, 2, 0, "", _
                                        "Descripción", 2, "Descripcion", 1800, 0, 0, "", _
                                        "Boton", 2, "Boton", 600, 0, 0, "", _
                                        "Activo", 2, "Activo", 600, 2, 4, "")
        Set gdrOrigenVenta.DataSource = RsOV
    'End If
    SSTab1.Tab = 0
    
    If pais = "003" Then
        Me.lblDescripcion.Visible = False
        Me.Label(8).Visible = False
        Me.txtEnlaceSunat.Visible = False
        Me.cmdBusca.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmTipoPedidoDetalle = Nothing
End Sub
 
Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoPedido.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoPedido.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoPedido.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoPedido.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoPedido.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoPedido.grdGrilla
    End Select
    Asignar
End Sub

Sub Asignar()
    
    With frmTipoPedido.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtTipoPedido = IIf(IsNull(!Descripcion), "", !Descripcion)
        chkActivo = IIf(IsNull(!Activo), 0, IIf(!Activo, 1, 0))
    End With
    Set RsDetalle = Lib.OpenRecordset("select tCodigoCanalVenta  ,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat, ISNULL(lObligaOrigenVenta,0)lObligaOrigenVenta , isnull(lMobileInforest,0)lMobileInforest, isnull(lActivaCuentasXCobrar,0)lActivaCuentasXCobrar from TCANALVENTA where tCodigoCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'", Cn)

    If Not (RsDetalle.EOF Or RsDetalle.BOF) Then
       With RsDetalle
            chkObligaMesa = IIf(IsNull(!lobligamesa), 0, IIf(!lobligamesa, 1, 0))
            chkObligaPax = IIf(IsNull(!lObligaPax), 0, IIf(!lObligaPax, 1, 0))
            chkActivaMozo = IIf(IsNull(!lactivamozo), 0, IIf(!lactivamozo, 1, 0))
            chkObligaMozo = IIf(IsNull(!lobligamozo), 0, IIf(!lobligamozo, 1, 0))
            chkEntregarA = IIf(IsNull(!lobligaentregara), 0, IIf(!lobligaentregara, 1, 0))
            chkActivaMotorizado = IIf(IsNull(!lactivamotorizado), 0, IIf(!lactivamotorizado, 1, 0))
            chkObligaMotorizado = IIf(IsNull(!lobligamotorizado), 0, IIf(!lobligamotorizado, 1, 0))
            chkCentralPedido = IIf(IsNull(!lcanalCentralPedidos), 0, IIf(!lcanalCentralPedidos, 1, 0))
            chkDelivery = IIf(IsNull(!lCanalDelivery), 0, IIf(!lCanalDelivery, 1, 0))
            chkFechaEntregaDelivery = IIf(IsNull(!lObligaIngresoFechaEntrega), 0, IIf(!lObligaIngresoFechaEntrega, 1, 0))
            chklCuentaXCobrar = IIf(IsNull(!lActivaCuentasXCobrar), 0, IIf(!lActivaCuentasXCobrar, 1, 0))
            Me.txtEnlaceContable1.Text = IIf(IsNull(!tenlacecontable1), "", !tenlacecontable1)
            Me.txtEnlaceContable2.Text = IIf(IsNull(!tenlacecontable2), "", !tenlacecontable2)
            chkObligaClienteFrecuente = IIf(IsNull(!lobligaclientefrecuente), 0, IIf(!lobligaclientefrecuente, 1, 0))
            Me.txtEnlaceSunat.Text = IIf(IsNull(!tCodigoSunat), "", !tCodigoSunat)
            ChklObligaOrigenVenta = IIf(IsNull(!lObligaOrigenVenta), 0, IIf(!lObligaOrigenVenta, 1, 0))
            lblDescripcion.Caption = Calcular("select isnull(descripcion,'') as codigo from tproductosunat where tcodigoproducto='" & Me.txtEnlaceSunat.Text & "'", Cn)
            chklMobileInforest = IIf(IsNull(!lMobileInforest), 0, IIf(!lMobileInforest, 1, 0))
        'Frame5.Visible = True
        Isql = "select CodOrigenVenta,Descripcion,Boton,Visible,Delivery,Activo, tCodAPIWEB,  lVisualizaDespachador from vorigenventa where CodCanalVenta='" & Me.txtCodigo.Text & "'" '"select * from vOrigenVenta where CodCanalVenta='" & frmTipoPedido.RsCabecera.Fields(0) & "'"
        Set RsOV = Lib.OpenRecordset(Isql, Cn)
        Call ConfGrilla(4, gdrOrigenVenta, "Codigo", 2, "CodOrigenVenta", 700, 2, 0, "", _
                                        "Descripción", 2, "Descripcion", 1800, 0, 0, "", _
                                        "Boton", 2, "Boton", 600, 0, 0, "", _
                                        "Activo", 2, "Activo", 600, 2, 4, "")
        Set gdrOrigenVenta.DataSource = RsOV
        
        End With
    Else
        chkObligaMesa = 0
        chkObligaPax = 0
        chkActivaMozo = 0
        chkObligaMozo = 0
        chkActivaMotorizado = 0
        chkObligaMotorizado = 0
        chkCentralPedido = 0
        chkDelivery = 0
        chkEntregarA = 0
        chkFechaEntregaDelivery = 0
        Me.txtEnlaceContable1.Text = ""
        Me.txtEnlaceContable2.Text = ""
        chkObligaClienteFrecuente = 0
        Me.txtEnlaceSunat.Text = ""
        lblDescripcion.Caption = ""
        ChklObligaOrigenVenta = 0
        chklMobileInforest = 0
    End If
    cmdTexto.Caption = "Registro " & IIf(frmTipoPedido.RsCabecera.RecordCount = 0, 0, frmTipoPedido.RsCabecera.AbsolutePosition) & " de " & frmTipoPedido.RsCabecera.RecordCount
End Sub


Private Sub txtbotonOV_Change()
    If Not IsNumeric(txtbotonOV.Text) Then
        'MsgBox "Ingrese un numero entero"
        txtbotonOV.Text = ""
        Exit Sub
    End If
    If CInt(txtbotonOV.Text) > 15 Then
        MsgBox "El valor maximo de ingreso de botonera es 15 "
        txtbotonOV.Text = ""
    End If

End Sub


Private Sub txtEnlaceContable1_LostFocus()
    Call ValidaStr(txtEnlaceContable1)
End Sub

 

Private Sub txtEnlaceContable2_LostFocus()
  Call ValidaStr(txtEnlaceContable2)
End Sub

Private Sub txtTipoPedido_LostFocus()
   Call ValidaStr(txtTipoPedido)
End Sub
