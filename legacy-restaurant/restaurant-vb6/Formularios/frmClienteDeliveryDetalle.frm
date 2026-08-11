VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClienteDeliveryDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7395
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11460
   Icon            =   "frmClienteDeliveryDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11460
   Begin TabDlg.SSTab SSTab1 
      Height          =   6555
      Left            =   1980
      TabIndex        =   43
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   11562
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "frmClienteDeliveryDetalle.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cuenta Corriente y Facturación"
      TabPicture(1)   =   "frmClienteDeliveryDetalle.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "chkCuentaCorriente"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "frmCtaCte"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Productos Asociados"
      TabPicture(2)   =   "frmClienteDeliveryDetalle.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Parientes"
      TabPicture(3)   =   "frmClienteDeliveryDetalle.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   87
         Top             =   360
         Width           =   9135
         Begin TrueOleDBGrid80.TDBGrid grdParientes 
            Height          =   5445
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   9604
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
      End
      Begin VB.Frame Frame4 
         Caption         =   " Otros Datos Comerciales "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   -74880
         TabIndex        =   73
         Top             =   2550
         Width           =   9135
         Begin VB.TextBox txtTarjeta 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2340
            MaxLength       =   10
            TabIndex        =   21
            Text            =   " "
            Top             =   690
            Width           =   2500
         End
         Begin VB.TextBox txtDescuento 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2340
            MaxLength       =   20
            TabIndex        =   22
            Text            =   " "
            Top             =   1080
            Width           =   2500
         End
         Begin VB.CheckBox chkPunto 
            Alignment       =   1  'Right Justify
            Caption         =   "Canjes por Puntos  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5655
            TabIndex        =   23
            Top             =   300
            Width           =   1800
         End
         Begin VB.TextBox txtPunto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Height          =   285
            Left            =   7275
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1350
         End
         Begin VB.TextBox txtPunto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Height          =   285
            Left            =   7275
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   615
            Width           =   1350
         End
         Begin VB.TextBox txtPunto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Height          =   285
            Left            =   7275
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   990
            Width           =   1350
         End
         Begin MSDataListLib.DataCombo cboTarjeta 
            Height          =   330
            Left            =   2340
            TabIndex        =   20
            Top             =   300
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Caption         =   "Número de Tarjeta Bancaria :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   120
            TabIndex        =   84
            Top             =   750
            Width           =   2100
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tarjeta Bancaria :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   945
            TabIndex        =   81
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descuento :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   1335
            TabIndex        =   80
            Top             =   1140
            Width           =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Puntos Disponibles :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   5730
            TabIndex        =   79
            Top             =   1425
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Puntos Acumulados :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   5670
            TabIndex        =   78
            Top             =   660
            Width           =   1530
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Puntos Utilizados :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   5865
            TabIndex        =   77
            Top             =   1035
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   63
         Top             =   360
         Width           =   9135
         Begin VB.CheckBox tCanal4 
            Height          =   255
            Left            =   3960
            TabIndex        =   117
            Top             =   5400
            Width           =   4125
         End
         Begin VB.CheckBox tCanal1 
            Height          =   255
            Left            =   3960
            TabIndex        =   116
            Top             =   4680
            Width           =   2320
         End
         Begin MSComCtl2.DTPicker tHoraFin 
            Height          =   375
            Left            =   2280
            TabIndex        =   115
            Top             =   5520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   56295427
            UpDown          =   -1  'True
            CurrentDate     =   43888.9993055556
         End
         Begin MSComCtl2.DTPicker tHoraInicio 
            Height          =   375
            Left            =   600
            TabIndex        =   114
            Top             =   5520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   56295427
            UpDown          =   -1  'True
            CurrentDate     =   43888
         End
         Begin VB.CheckBox tCanal5 
            Height          =   255
            Left            =   3960
            TabIndex        =   113
            Top             =   5640
            Width           =   4125
         End
         Begin VB.CheckBox tCanal3 
            Height          =   255
            Left            =   3960
            TabIndex        =   110
            Top             =   5160
            Width           =   2320
         End
         Begin VB.CheckBox tCanal2 
            Height          =   255
            Left            =   3960
            TabIndex        =   109
            Top             =   4920
            Width           =   2320
         End
         Begin VB.CheckBox lSoloxDia 
            Caption         =   "Aplica solo una vez por día"
            Height          =   255
            Left            =   6480
            TabIndex        =   108
            Top             =   5040
            Width           =   2535
         End
         Begin VB.CheckBox lRangoHorarios 
            Caption         =   "Aplicar Rango de Horas:"
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
            Left            =   120
            TabIndex        =   107
            Top             =   5160
            Width           =   3375
         End
         Begin VB.CheckBox chkExcluyeProductos 
            Caption         =   "Excluir Productos no Asociados"
            Height          =   195
            Left            =   6480
            TabIndex        =   29
            ToolTipText     =   "Si activa esta opción solo se podrá seleccionar productos asociados a este cliente"
            Top             =   4800
            Width           =   2580
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
            Height          =   495
            Index           =   6
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Quitar "
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
            Index           =   9
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   720
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
            Left            =   1620
            MaxLength       =   15
            TabIndex        =   68
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4740
            Width           =   1680
         End
         Begin VB.Frame Frame3 
            Caption         =   " Seleccionar Producto de Venta "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   7515
            Begin VB.CommandButton CmdBuscaProducto 
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
               Left            =   6360
               Picture         =   "frmClienteDeliveryDetalle.frx":04B2
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   240
               Width           =   705
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
               Left            =   960
               MaxLength       =   10
               TabIndex        =   27
               Text            =   "0.00"
               Top             =   660
               Width           =   1020
            End
            Begin VB.CheckBox chkExcluyente 
               Caption         =   "Permite Otros Descuentos"
               Height          =   195
               Left            =   2400
               TabIndex        =   28
               Top             =   705
               Width           =   2250
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Producto :"
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   67
               Top             =   300
               Width           =   735
            End
            Begin VB.Label txtProducto 
               BackColor       =   &H8000000E&
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
               Height          =   300
               Left            =   960
               TabIndex        =   66
               Top             =   240
               Width           =   5280
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Precio :"
               Height          =   195
               Index           =   22
               Left            =   315
               TabIndex        =   65
               Top             =   705
               Width           =   540
            End
         End
         Begin TrueOleDBGrid80.TDBGrid grdGrilla 
            Height          =   3075
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   5424
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
         Begin VB.Label Label3 
            Caption         =   "Aplica a Canal de Venta?"
            Height          =   255
            Left            =   3960
            TabIndex        =   118
            Top             =   4440
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Fin:"
            Height          =   255
            Left            =   1920
            TabIndex        =   112
            Top             =   5580
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Inicio:"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   5580
            Width           =   495
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total de Productos : "
            Height          =   195
            Index           =   24
            Left            =   60
            TabIndex        =   71
            Top             =   4785
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox chkCuentaCorriente 
         Caption         =   "Habilitar Cuenta Corriente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   62
         Top             =   420
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "Facturación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   -74880
         TabIndex        =   61
         Top             =   4485
         Width           =   9135
         Begin VB.Frame Frame6 
            Caption         =   "Facturacion a Nombre de"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   90
            TabIndex        =   96
            Top             =   645
            Width           =   8955
            Begin VB.TextBox txtRazonSocial 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   2265
               MaxLength       =   80
               TabIndex        =   100
               Text            =   " "
               Top             =   750
               Width           =   6270
            End
            Begin VB.TextBox txtRuc 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   315
               Width           =   2500
            End
            Begin VB.CommandButton cmdLimpiar 
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
               Left            =   5400
               Picture         =   "frmClienteDeliveryDetalle.frx":05B4
               Style           =   1  'Graphical
               TabIndex        =   98
               ToolTipText     =   "Buscar Razón Social"
               Top             =   330
               Width           =   345
            End
            Begin VB.CommandButton cmdBuscaCliente 
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
               Left            =   4800
               Picture         =   "frmClienteDeliveryDetalle.frx":06B6
               Style           =   1  'Graphical
               TabIndex        =   97
               Top             =   330
               Width           =   570
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Razón Social :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1065
               TabIndex        =   102
               Top             =   810
               Width           =   1035
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Identificador Tributario :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   435
               TabIndex        =   101
               Top             =   390
               Width           =   1695
            End
         End
         Begin VB.CheckBox chkEmisionMonedaExtranjera 
            Caption         =   "Permite Emisión en Moneda Extranjera."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   330
            Width           =   3450
         End
      End
      Begin VB.Frame frmCtaCte 
         Caption         =   "   "
         Height          =   2055
         Left            =   -74880
         TabIndex        =   57
         Top             =   420
         Width           =   9135
         Begin VB.TextBox txtDiasCredito 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2775
            MaxLength       =   15
            TabIndex        =   104
            Text            =   " "
            Top             =   1650
            Width           =   2025
         End
         Begin VB.TextBox txtConsumoCobrar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Height          =   285
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1350
         End
         Begin VB.TextBox txtConsumoCte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Height          =   285
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   840
            Width           =   1350
         End
         Begin VB.TextBox txtCuentaCobrar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2775
            MaxLength       =   15
            TabIndex        =   19
            Text            =   " "
            Top             =   1275
            Width           =   2025
         End
         Begin VB.TextBox txtLinea 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2775
            MaxLength       =   15
            TabIndex        =   18
            Text            =   " "
            Top             =   900
            Width           =   2025
         End
         Begin MSDataListLib.DataCombo cboTipoCtaCte 
            Height          =   330
            Left            =   2295
            TabIndex        =   16
            Top             =   270
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboSubTipoCtaCte 
            Height          =   330
            Left            =   6495
            TabIndex        =   17
            Top             =   300
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Caption         =   "Visualización en columna ""Vencimiento"", modulo Ctas x Cobrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   33
            Left            =   4875
            TabIndex        =   105
            Top             =   1710
            Width           =   3930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Limite de Credito (Dias) :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   32
            Left            =   990
            TabIndex        =   103
            Top             =   1680
            Width           =   1755
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Consumo en Cuenta Corriente :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   31
            Left            =   5130
            TabIndex        =   92
            Top             =   960
            Width           =   2250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Consumo en Cuenta Por Cobrar :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   30
            Left            =   5010
            TabIndex        =   91
            Top             =   1320
            Width           =   2370
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Por Cobrar :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   29
            Left            =   1320
            TabIndex        =   90
            Top             =   1320
            Width           =   1425
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   28
            Left            =   1440
            TabIndex        =   89
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Linea de Crédito :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   960
            TabIndex        =   60
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cta.Cte. :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   20
            Left            =   1185
            TabIndex        =   59
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "SubTipo Cta.Cte. :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   19
            Left            =   5100
            TabIndex        =   58
            Top             =   360
            Width           =   1290
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
         Height          =   6100
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   9195
         Begin VB.CommandButton BtnBuscarCentral 
            Caption         =   "Buscar"
            Height          =   370
            Left            =   4920
            TabIndex        =   106
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarFoto 
            Caption         =   "Editar"
            Height          =   350
            Left            =   6600
            TabIndex        =   86
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtIdentidad 
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
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   4
            Text            =   " "
            Top             =   2200
            Width           =   2610
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   360
            Width           =   1170
         End
         Begin VB.TextBox txtEmail 
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
            MaxLength       =   50
            TabIndex        =   12
            Text            =   " "
            Top             =   4950
            Width           =   6780
         End
         Begin VB.TextBox txtTelefono 
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
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   2
            Text            =   " "
            Top             =   1170
            Width           =   2610
         End
         Begin VB.TextBox txtDireccion 
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
            MaxLength       =   100
            TabIndex        =   7
            Text            =   " "
            Top             =   3375
            Width           =   6780
         End
         Begin VB.TextBox txtReferencia 
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
            MaxLength       =   220
            TabIndex        =   8
            Text            =   " "
            Top             =   3765
            Width           =   6780
         End
         Begin VB.CheckBox chkActivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo :"
            Height          =   195
            Left            =   1560
            TabIndex        =   15
            Top             =   5820
            Width           =   900
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
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   6
            Text            =   " "
            Top             =   2985
            Width           =   6780
         End
         Begin VB.TextBox txtApellido 
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
            MaxLength       =   50
            TabIndex        =   5
            Text            =   " "
            Top             =   2595
            Width           =   6780
         End
         Begin VB.TextBox txtObservacion 
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
            MaxLength       =   250
            TabIndex        =   11
            Text            =   " "
            Top             =   4560
            Width           =   6780
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   2280
            TabIndex        =   13
            Top             =   5340
            Width           =   2610
            _ExtentX        =   4604
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd MMMM"
            Format          =   56295427
            UpDown          =   -1  'True
            CurrentDate     =   38891
         End
         Begin MSDataListLib.DataCombo cboZona 
            Height          =   315
            Left            =   6450
            TabIndex        =   10
            Top             =   4140
            Width           =   2610
            _ExtentX        =   4604
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
         Begin MSDataListLib.DataCombo cboTipoCliente 
            Height          =   315
            Left            =   2280
            TabIndex        =   1
            Top             =   750
            Width           =   2610
            _ExtentX        =   4604
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
         Begin MSDataListLib.DataCombo cboDistrito 
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Top             =   4140
            Width           =   2610
            _ExtentX        =   4604
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
         Begin MSDataListLib.DataCombo cboEstadoCliente 
            Height          =   315
            Left            =   6480
            TabIndex        =   14
            Top             =   5340
            Width           =   2610
            _ExtentX        =   4604
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
         Begin MSDataListLib.DataCombo cboTipoIdentidad 
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Top             =   1815
            Width           =   2610
            _ExtentX        =   4604
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
         Begin MSComDlg.CommonDialog dlgFoto 
            Left            =   8400
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   2055
            Left            =   6240
            Stretch         =   -1  'True
            ToolTipText     =   "Foto"
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Teléfono :"
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
            Height          =   195
            Index           =   27
            Left            =   1335
            TabIndex        =   85
            Top             =   1230
            Width           =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Número Identidad :"
            Height          =   195
            Index           =   25
            Left            =   840
            TabIndex        =   83
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Identidad :"
            Height          =   195
            Index           =   16
            Left            =   870
            TabIndex        =   82
            Top             =   1875
            Width           =   1335
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Index           =   0
            Left            =   1620
            TabIndex        =   72
            Top             =   405
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail :"
            Height          =   195
            Index           =   6
            Left            =   1680
            TabIndex        =   56
            Top             =   4995
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Dirección :"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   55
            Top             =   3420
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   2265
            TabIndex        =   54
            Top             =   1470
            Width           =   2625
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Referencia :"
            Height          =   195
            Index           =   3
            Left            =   1335
            TabIndex        =   53
            Top             =   3810
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Razón Social / Nombres :"
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   52
            Top             =   3030
            Width           =   1830
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Razón Comercial / Apellido :"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   51
            Top             =   2640
            Width           =   2010
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Observación :"
            Height          =   195
            Index           =   8
            Left            =   1215
            TabIndex        =   50
            Top             =   4605
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Nacimiento :"
            Height          =   195
            Index           =   9
            Left            =   600
            TabIndex        =   49
            Top             =   5400
            Width           =   1605
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Zona :"
            Height          =   195
            Index           =   10
            Left            =   5925
            TabIndex        =   48
            Top             =   4200
            Width           =   465
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cliente Frecuente :"
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   47
            Top             =   810
            Width           =   1695
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Index           =   17
            Left            =   1635
            TabIndex        =   46
            Top             =   4200
            Width           =   570
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            Height          =   195
            Index           =   18
            Left            =   5820
            TabIndex        =   45
            Top             =   5400
            Width           =   585
         End
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   11400
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6615
      Width           =   11460
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Tiendas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   7740
         Picture         =   "frmClienteDeliveryDetalle.frx":07B8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   10080
         Picture         =   "frmClienteDeliveryDetalle.frx":08B2
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   60
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   8910
         Picture         =   "frmClienteDeliveryDetalle.frx":09A4
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5400
         Picture         =   "frmClienteDeliveryDetalle.frx":0AA6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   6570
         Picture         =   "frmClienteDeliveryDetalle.frx":0FD8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5250
         TabIndex        =   41
         Top             =   60
         Width           =   5310
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   470
            Picture         =   "frmClienteDeliveryDetalle.frx":150A
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   940
            Picture         =   "frmClienteDeliveryDetalle.frx":1A4C
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmClienteDeliveryDetalle.frx":1F8E
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4785
            Picture         =   "frmClienteDeliveryDetalle.frx":24D0
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4305
            Picture         =   "frmClienteDeliveryDetalle.frx":2A12
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3840
            Picture         =   "frmClienteDeliveryDetalle.frx":2F54
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   470
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   42
            Top             =   150
            Width           =   2370
         End
      End
   End
   Begin VB.Image Image 
      Height          =   6465
      Left            =   0
      Picture         =   "frmClienteDeliveryDetalle.frx":3496
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "frmClienteDeliveryDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsZona As Recordset
Dim RsTipoCliente As Recordset
Dim RsTarjeta As Recordset
Dim RsDistrito As Recordset
Dim nZona As Double
Dim nCodigo As String
Dim sTelefono As String
Dim nDescuento As Double
Dim RsEstadoFrecuente As Recordset
Dim RsTipoIdentidad As Recordset

Dim sClienteFactura As String

Dim RsTipoCtaCte As Recordset
Dim RsSubTipoCtaCte As Recordset
Dim nLinea As Double
Dim nLineaXCobrar As Double
Dim RsProducto As Recordset
  
Dim sTemporal As String
Dim nColumna As Integer
Dim nPrecio As Double

Dim i As Integer

Dim validaTipoIdentidad As Boolean
Private strFilenameRuta As String

Dim nDiasCredito As Integer

Dim sTablaParientes As String
Dim rsParientes As Recordset
Dim xrsParientes As Recordset


Sub LlenaCombos()
    With cboDistrito
         Isql = "Select * from vDistrito order by Descripcion"
         Set RsDistrito = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsDistrito
             .DataField = "descripcion"
             .ListField = "descripcion"
             .BoundColumn = "Codigo"
    End With

    With cboZona
         Isql = "Select * from vZona order by Descripcion"
         Set RsZona = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsZona
             .DataField = "descripcion"
             .ListField = "descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoCliente
        Isql = "Select * from vTipoClienteFrecuente Where Codigo<>'00' order by Codigo"
        Set RsTipoCliente = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsTipoCliente
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With
    
    With cboTarjeta
         Isql = "select tCodigoTarjeta as Codigo, tDetallado as Descripcion from tTarjetaCredito where lActivo=1"
         Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTarjeta
             .DataField = "descripcion"
             .ListField = "descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboEstadoCliente
         Isql = "select  Codigo,  tresumido as Descripcion from vEstadoFrecuente where lActivo=1"
         Set RsEstadoFrecuente = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstadoFrecuente
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoCtaCte
         Isql = "Select * from vTipoCtaCte where lActivo=1 order by Descripcion"
         Set RsTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoCtaCte
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboSubTipoCtaCte
         Isql = "Select * from vSubTipoCtaCte where lActivo=1 order by Descripcion"
         Set RsSubTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipoCtaCte
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoIdentidad
         Isql = "Select * from vTipoIdentidad where lActivo=1 order by Descripcion"
         Set RsTipoIdentidad = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIdentidad
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    
End Sub

Sub Asignar()
    With frmClienteDelivery.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtTarjeta = IIf(IsNull(!tNumeroTarjeta), "", !tNumeroTarjeta)
        txtApellido = IIf(IsNull(!tApellido), "", !tApellido)
        txtNombre = IIf(IsNull(!tNombre), "", !tNombre)
        txtTelefono = IIf(IsNull(!tTelefono), "", !tTelefono)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtReferencia = IIf(IsNull(!tReferencia), "", !tReferencia)
        txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
        txtEmail = IIf(IsNull(!temail), "", !temail)
        nDescuento = IIf(IsNull(!nDescuento), 0, !nDescuento)
        cboTipoCliente.BoundText = IIf(IsNull(!tTipoCliente), "00", !tTipoCliente)
        cboDistrito.BoundText = IIf(IsNull(!tDistrito), "", !tDistrito)
        cboTarjeta.BoundText = IIf(IsNull(!tCodigoTarjeta), "", !tCodigoTarjeta)
        txtDescuento = Format(nDescuento, "##0.00")
        txtPunto.Text = Format(IIf(IsNull(!nDisponible), 0, !nDisponible), "###,##0.00")
        txtPunto1.Text = Format(IIf(IsNull(!nAcumulado), 0, !nAcumulado), "###,##0.00")
        txtPunto2.Text = Format(IIf(IsNull(!nUtilizado), 0, !nUtilizado), "###,##0.00")
                    
        'Data Combo
        cboZona.BoundText = IIf(IsNull(!tZona), "", Trim(!tZona))
        cboEstadoCliente.BoundText = IIf(IsNull(!testadofrecuente), "", Trim(!testadofrecuente))
        
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
                          
        txtIdentidad.Text = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        cboTipoIdentidad.BoundText = IIf(IsNull(!tTipoIdentidad), "", !tTipoIdentidad)
        
        
        chkPunto = IIf(!lPuntos = True, 1, 0)
        dtpFecha.Value = IIf(IsNull(!fNacimiento), Null, !fNacimiento)
        If IsNull(!lexcluyeproductos) Then
            chkExcluyeProductos = 0
        Else
            If !lexcluyeproductos Then
                chkExcluyeProductos = 1
            Else
                chkExcluyeProductos = 0
            End If
        End If
        
        nLinea = IIf(IsNull(!nLinea), 0, !nLinea)
        txtLinea = Format(nLinea, "######0.00")
                                  
                                  
        nLineaXCobrar = IIf(IsNull(!nLineaporcobrar), 0, !nLineaporcobrar)
        txtCuentaCobrar = Format(nLineaXCobrar, "######0.00")
        
        Me.txtConsumoCte.Text = IIf(IsNull(!nCOnsumo), 0, !nCOnsumo)
        Me.txtConsumoCobrar.Text = IIf(IsNull(!nCOnsumoporcobrar), 0, !nCOnsumoporcobrar)
        
        cboTipoCtaCte.BoundText = IIf(IsNull(!tTipoCtaCte), "", !tTipoCtaCte)
        cboSubTipoCtaCte.BoundText = IIf(IsNull(!tSubTipoCtaCte), "", !tSubTipoCtaCte)
                                 
        If IsNull(!lClienteCtaCte) Then
            chkCuentaCorriente = 0
            txtLinea.Enabled = False
            nLinea = 0
            txtLinea.Text = "0.00"
            cboTipoCtaCte.Enabled = False
            cboSubTipoCtaCte.Enabled = False
            cboTipoCtaCte.Text = ""
            cboSubTipoCtaCte.Text = ""
        Else
            chkCuentaCorriente = IIf(!lClienteCtaCte = True, 1, 0)
        End If
        chkCuentaCorriente_Click
        
        sClienteFactura = IIf(IsNull(!tCodigoCliente), "", !tCodigoCliente)
        If sClienteFactura = "" Then
            txtRuc = ""
            txtRazonSocial = ""
        Else
            txtRuc = Calcular("select isnull(tidentidad,'') as codigo from tcliente where tcodigocliente='" & sClienteFactura & "'", Cn)
            Me.txtRazonSocial = Calcular("select isnull(tempresa,'') as codigo from tcliente where tcodigocliente='" & sClienteFactura & "'", Cn)
        End If
        chkEmisionMonedaExtranjera = IIf(!lEmisionMonedaExtranjera = True, 1, 0)
        
        nDiasCredito = IIf(IsNull(!nDiasCredito), 0, !nDiasCredito)
        txtDiasCredito = Format(nDiasCredito, "######0")
        'GCAA
        lRangoHorarios.Value = IIf(!lRangoHorarios = True, 1, 0)
        tHoraInicio = Format(!tHoraInicio.Value, "hh:mm")
        tHoraFin = Format(!tHoraFin.Value, "hh:mm")
        tCanal1.Value = IIf(!tCanal1 = True, 1, 0)
        tCanal2.Value = IIf(!tCanal2 = True, 1, 0)
        tCanal3.Value = IIf(!tCanal3 = True, 1, 0)
        tCanal4.Value = IIf(!tCanal4 = True, 1, 0)
        tCanal5.Value = IIf(!tCanal5 = True, 1, 0)
        lSoloxDia.Value = IIf(!lSoloxDia = True, 1, 0)
        
        
        Isql = "select  tDetallado , tCodigoCanalVenta from TCANALVENTA order by tCodigoCanalVenta asc"
        Set RsxCanalventa = Lib.OpenRecordset(Isql, Cn)
        
        RsxCanalventa.MoveFirst
        Do Until RsxCanalventa.EOF
             If RsxCanalventa!tCodigoCanalVenta = "01" Then
                tCanal1.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "02" Then
                tCanal2.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "03" Then
                tCanal3.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "04" Then
                tCanal4.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "05" Then
                tCanal5.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado
             End If
             
            RsxCanalventa.MoveNext
        Loop
        
        
        
        
    End With
       Cn.Execute "delete from " & sTemporal
       Cn.Execute "insert into " & sTemporal & " (tcodigodelivery, tCodigoProducto,nprecio,lPermiteDescuento) select tcodigodelivery, tCodigoProducto,nPrecio,lPermiteDescuentos from tclienteproducto where tcodigodelivery='" & txtCodigo.Text & "'"
       RsProducto.Requery
       txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
       
       Cn.Execute "delete from " & sTablaParientes
       Cn.Execute "insert into " & sTablaParientes & " (Codigo,Pariente,Conyugue, Hijo) SELECT tcodigoPariente, pariente, lconyugue,lhijo from vPariente where tcodigodelivery='" & txtCodigo.Text & "' order by 2 asc"
       rsParientes.Requery
       
       'fotoooooooooo
       On Error GoTo err
       Dim rst1 As New ADODB.Recordset
       imgFoto.DataField = "foto"
       cmdAgregarFoto.Caption = "Editar"
       Set rst1 = Lib.OpenRecordset("select ifoto as foto from tdelivery where tcodigodelivery='" & txtCodigo.Text & "'", Cn)
       Set imgFoto.DataSource = rst1
       Exit Sub
       
err:
    imgFoto.Picture = Nothing
End Sub
Private Sub BtnBuscarCentral_Click()
 
  
  If Trim(txtIdentidad.Text) = "" Then
    MsgBox "Debe Ingresar Numero de Identidad para poder realizar la Busqueda", vbExclamation, sMensaje
    Exit Sub
  End If
  
  If VerConexionInfoRestCentral = "ON" Then
    If Calcular("select count(tidentidad) as Codigo from TDELIVERY where tidentidad='" & Trim(txtIdentidad.Text) & "'", Conex) > 0 Then
       MsgBox "Cliente si Existe en Base Central", vbExclamation, sMensaje
       Isql = "select isnull(TipoCliente,'')as  TipoCliente, tTelefono, isnull(tTipoIdentidad,'')as tTipoIdentidad, tApellido,tNombre,tDireccion,tReferencia,ISNULL(tDistrito,'') AS tDistrito,ISNULL(tZona,'')AS tZona ,tObservacion,tEmail,fNacimiento,lActivo From vDelivery Where tidentidad='" & Trim(txtIdentidad.Text) & "'"
        
       Set xrsParientes = Lib.OpenRecordset(Isql, Conex)
         
        cboTipoCliente.BoundText = xrsParientes!TipoCliente
        txtTelefono.Text = xrsParientes!tTelefono
        txtApellido.Text = xrsParientes!tApellido
        txtNombre.Text = xrsParientes!tNombre
        txtDireccion.Text = xrsParientes!tDireccion
       txtReferencia.Text = xrsParientes!tReferencia
       txtObservacion.Text = xrsParientes!tObservacion
       txtEmail.Text = xrsParientes!temail
    Else
        MsgBox "Cliente No Existe", vbExclamation, sMensaje
    End If

  End If





End Sub

 

Private Sub chkCuentaCorriente_Click()
    If chkCuentaCorriente Then
        txtLinea.Enabled = True
        txtCuentaCobrar.Enabled = True
        cboTipoCtaCte.Enabled = True
        cboSubTipoCtaCte.Enabled = True
        txtDiasCredito.Enabled = True
    Else
        txtLinea.Enabled = False
        nLinea = 0
        nLineaXCobrar = 0
        txtCuentaCobrar.Enabled = False
        txtCuentaCobrar.Text = "0.00"
        txtLinea.Text = "0.00"
        cboTipoCtaCte.Enabled = False
        cboSubTipoCtaCte.Enabled = False
        cboTipoCtaCte.Text = ""
        cboSubTipoCtaCte.Text = ""
        txtDiasCredito.Text = "0"
        txtDiasCredito.Enabled = False
    End If
End Sub

Private Sub cmdAgregarFoto_Click()

'on error GoTo ErrHandler
If txtCodigo.Text <> "" Then
    dlgFoto.CancelError = False
    With cmdAgregarFoto
        If .Caption = "Editar" Then
            dlgFoto.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFoto.FileName = ""
            dlgFoto.ShowOpen
            imgFoto.Visible = True
            If dlgFoto.FileName <> "" Then
                .Caption = "Guardar"
                strFilenameRuta = dlgFoto.FileName
                imgFoto.Picture = LoadPicture(strFilenameRuta)
            End If
        Else
            
            Call GuardarFoto
            .Caption = "Editar"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para el Cliente"
    Exit Sub
End If
errHandler:
strFilenameRuta = ""
    Exit Sub
End Sub

Public Sub GuardarFoto()
        
        Dim imgTeacher()      As Byte
        Dim varPhoto          As Variant
        Dim numfile           As Long
        If (strFilenameRuta <> "") Then
            varPhoto = FileLen(strFilenameRuta)
            ReDim bufimages(varPhoto - 1) As Byte
            numfile = FreeFile
            Open strFilenameRuta For Binary As #numfile
            Get #numfile, , bufimages
            Close #numfile
             imgTeacher = bufimages
        End If
        If (strFilenameRuta = "") Then
            imgTeacher = LoadResData(101, "CUSTOM")
            varPhoto = UBound(imgTeacher)
        End If
        Dim lnfoto As Variant
        lnfoto = varPhoto
        Dim Cmd As New ADODB.Command
        Dim prm As New ADODB.Parameter
        With Cmd
                .ActiveConnection = Cn
                .CommandText = "sp_UpdFotoDelivery"
                .CommandType = adCmdStoredProc
        End With
        Set prm = Cmd.CreateParameter("@tCodigo", adChar, adParamInput, 10, txtCodigo.Text)
        Cmd.Parameters.Append prm
        Set prm = Cmd.CreateParameter("@oFoto", adLongVarBinary, adParamInput, lnfoto + 1)
        Cmd.Parameters.Append prm
        If Not IsNull(imgTeacher) Then
            prm.AppendChunk imgTeacher
        Else
            prm.Value = Null
        End If
        Cmd.Execute
End Sub

Private Sub cmdBuscaCliente_Click()
        Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
        frmBusquedaRapida.cmdOpcion(1).Enabled = True
        frmBusquedaRapida.cmdOpcion(2).Enabled = True
        frmBusquedaRapida.cmdOpcion(3).Enabled = False
        frmBusquedaRapida.cmdOpcion(0).Enabled = False
        frmBusquedaRapida.nPredeterm = 1
        
        Select Case pais ' ok
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                        "NIT", 2, "tIdentidad", 1600, 2, 0, "", _
                                        "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                        "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                        "Cliente", 2, "Descripcion", 5500, 0, 0, "")
        End Select
        
        sTemp = Trim(txtRuc.Text)
        frmBusquedaRapida.Show vbModal
        
        If wEnter = True And sCodigo <> "" Then
                sClienteFactura = sCodigo
                txtRuc = Calcular("select isnull(tidentidad,'') as codigo from tcliente where tcodigocliente='" & sClienteFactura & "'", Cn)
                Me.txtRazonSocial = Calcular("select isnull(tempresa,'') as codigo from tcliente where tcodigocliente='" & sClienteFactura & "'", Cn)
            Else
                sClienteFactura = ""
                txtRuc = ""
                Me.txtRazonSocial = ""
                Exit Sub
        End If
End Sub

Private Sub cmdFoto_Click()

End Sub

Private Sub cmdLimpiar_Click()
    sClienteFactura = ""
    txtRuc = ""
    txtRazonSocial = ""
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmClienteDelivery.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmClienteDelivery.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmClienteDelivery.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmClienteDelivery.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmClienteDelivery.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmClienteDelivery.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmClienteDelivery.RsCabecera.RecordCount = 0, 0, frmClienteDelivery.RsCabecera.AbsolutePosition) & " de " & frmClienteDelivery.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
 Dim xtTipoIdentidad As String
   Select Case Index
          Case Is = 0 ' Agregar
                Sw = True
                ActivarBotones (False)
                Blanquear Me
                dtpFecha.Value = Null
                chkActivo.Value = 1
                txtDescuento.Text = "0.00"
                cboEstadoCliente.BoundText = "01"
                
                sClienteFactura = ""
                chkEmisionMonedaExtranjera.Value = 0
                
                'Cambia el Nombre del Primer Text
                txtTarjeta.SetFocus
                Cn.Execute "delete from " & sTemporal
                Cn.Execute "insert into " & sTemporal & " (tcodigodelivery, tCodigoProducto,nprecio,lPermiteDescuento) select tcodigodelivery, tCodigoProducto,nPrecio,lPermiteDescuentos from tclienteproducto where tcodigodelivery='" & txtCodigo.Text & "'"
                RsProducto.Requery
                
                Cn.Execute "delete from " & sTablaParientes
                Cn.Execute "insert into " & sTablaParientes & " (Codigo,Pariente,Conyugue, Hijo) SELECT tcodigoPariente, pariente, lconyugue,lhijo from vPariente where tcodigodelivery='" & txtCodigo.Text & "' order by 2 asc"
                rsParientes.Requery
       
                cmdAgregarFoto.Caption = "Editar"
                imgFoto.Picture = Nothing
                txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                
          
          Case Is = 1 ' Grabar
          
'          If tHoraFin.Value <> "12:00:00 a.m." Or tHoraFin.Value <> "00:00:00" Then
             If tHoraInicio.Value > tHoraFin.Value Then
                  tHoraInicio.Value = tHoraFin.Value
                  MsgBox "Error en Horas", vbExclamation, sMensaje
                  Exit Sub
            '      MsgBox tHoraFin.Value, vbExclamation, sMensaje
               End If
            
'            End If

               Dim nCorrela As String
               Dim xnCorrela As String
               Dim NuevoCorrelativo As String
               Dim tValorDelivery As String
               
                  '-----Validacion de Campos----
                  If cboTipoCliente.BoundText = "" Then MsgBox ("!No ha seleccionado Tipo de Cliente¡"): cboTipoCliente.SetFocus: Exit Sub
'                  If cboTipoIdentidad.BoundText = "" Then MsgBox ("!No ha seleccionado Tipo de Documento¡"): Exit Sub
'                  If txtIdentidad.Text = "" Then MsgBox ("!Ingrese Nro. Documento de Identidad¡"): Exit Sub
'                  If txtApellido.Text = "" Then MsgBox ("!Ingrese Razón Comercial/Apellido¡"): Exit Sub
                  '------------------------
                  
               tValorDelivery = Calcular("Select tValor As Codigo from vTipoClienteFrecuente Where Codigo='" & Trim(cboTipoCliente.BoundText) & "'", Cn)
                                
                                
             If pais = "002" Then ' ECUADOR
                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & cboTipoIdentidad.BoundText & "'", Cn) Then
                            If Len(Trim(txtIdentidad.Text)) = 13 Or Len(Trim(txtIdentidad.Text)) = 10 Then
        
                            Else
                               MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
                               Exit Sub
                            End If
                        End If
                        
                        If Len(Trim(txtIdentidad.Text)) = 10 Then
                            xtTipoIdentidad = "01"
                        ElseIf Len(Trim(txtIdentidad.Text)) = 13 Then
                            xtTipoIdentidad = "02"
                        End If
                        
                        
                Else
                        'PERU - BOLIVIA
                        If Me.cboTipoIdentidad.BoundText = "02" Then
                        
                                If lLongitud And Len(Trim(txtIdentidad.Text)) <> nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                ElseIf Not lLongitud And Len(Trim(txtIdentidad.Text)) < nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                End If
                                
                                validaTipoIdentidad = False
                                validaTipoIdentidad = Calcular("select isnull(nvalor,0) as codigo from vTipoidentidad where Codigo='" & cboTipoIdentidad.BoundText & "' ", Cn)
                                If validaTipoIdentidad = True Then

                                
                                        If Not ValidaRuc(txtIdentidad.Text) Then
                                           MsgBox "El número  Identificador ingresado no es válido", vbCritical, sMensaje
                                           Exit Sub
                                        End If
                                End If
                        End If
                        xtTipoIdentidad = ""
               End If
                                               
                                
                                
               'Chequea Datos
               If Sw Then

            If VerConexionInfoRestCentral = "ON" Then
                If Calcular("select count(tIdentidad) as Codigo from TDELIVERY where tIdentidad='" & Trim(txtIdentidad.Text) & "'", Cn) > 0 Then
                   MsgBox "DNI Existente", vbExclamation, sMensaje
                   Exit Sub
                End If
            Else
                  If tValorDelivery = "1" Then
                        If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                           MsgBox "Teléfono Existente", vbExclamation, sMensaje
                           Exit Sub
                        End If
                  End If
            End If
                
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigodelivery) as Codigo from TDELIVERY", Cn)
                  If Mid(nCorrela, 1, 1) = "0" Or Mid(nCorrela, 1, 1) = "1" Or Mid(nCorrela, 1, 1) = "2" Or Mid(nCorrela, 1, 1) = "3" Or Mid(nCorrela, 1, 1) = "4" Or Mid(nCorrela, 1, 1) = "5" Or Mid(nCorrela, 1, 1) = "6" Or Mid(nCorrela, 1, 1) = "7" Or Mid(nCorrela, 1, 1) = "8" Or Mid(nCorrela, 1, 1) = "9" Then
                            If IsNull(nCorrela) Or nCorrela = "" Then
                                txtCodigo.Text = "0000001"
                            Else
                                txtCodigo.Text = Lib.Correlativo(nCorrela, 7)
                            End If
                   Else
                            If IsNull(nCorrela) Or nCorrela = "" Then
                                txtCodigo.Text = "0000001"
                            Else
                                If Mid(nCorrela, 2, 6) = "999999" Then
                                txtCodigo.Text = Chr(Asc(Mid(nCorrela, 1, 1)) + 1) + "000001"
                                Else
                                txtCodigo.Text = Mid(nCorrela, 1, 1) + Lib.Correlativo(Mid(nCorrela, 2, 6), 6)
                                End If
                   
                            End If
                   End If
                  Sw = False
                   
                   '=====================GCAA 26/02/2020 GRABAR CLIENTE
                   'Obtiene el Numero de Orden
                  
                  If VerConexionInfoRestCentral = "ON" Then
                        If Calcular("select count(tIdentidad) as Codigo from TDELIVERY where tIdentidad='" & Trim(txtIdentidad.Text) & "'", Conex) > 0 Then
                               
                           Else
                                xnCorrela = Calcular("select max(tCodigodelivery) as Codigo from TDELIVERY", Conex)
                                If Mid(xnCorrela, 1, 1) = "0" Or Mid(xnCorrela, 1, 1) = "1" Or Mid(xnCorrela, 1, 1) = "2" Or Mid(xnCorrela, 1, 1) = "3" Or Mid(xnCorrela, 1, 1) = "4" Or Mid(xnCorrela, 1, 1) = "5" Or Mid(xnCorrela, 1, 1) = "6" Or Mid(xnCorrela, 1, 1) = "7" Or Mid(xnCorrela, 1, 1) = "8" Or Mid(xnCorrela, 1, 1) = "9" Then
                                          If IsNull(xnCorrela) Or xnCorrela = "" Then
                                              NuevoCorrelativo = "0000001"
                                          Else
                                              NuevoCorrelativo = Lib.Correlativo(xnCorrela, 7)
                                          End If
                                 Else
                                          If IsNull(xnCorrela) Or xnCorrela = "" Then
                                              NuevoCorrelativo = "0000001"
                                          Else
                                              If Mid(xnCorrela, 2, 6) = "999999" Then
                                              NuevoCorrelativo = Chr(Asc(Mid(xnCorrela, 1, 1)) + 1) + "000001"
                                              Else
                                              NuevoCorrelativo = Mid(xnCorrela, 1, 1) + Lib.Correlativo(Mid(xnCorrela, 2, 6), 6)
                                              End If
                                 
                                          End If
                                 End If
                                Sw = False
                                 
                                  Isql = "insert into TDELIVERY( " & _
                                   "tCodigodelivery,  tCodigoTarjeta, tNumeroTarjeta, tApellido, tNombre, tDireccion, tTelefono, tReferencia, tZona, nDescuento, fNacimiento, tEMail, tObservacion, lActivo, lPuntos, nAcumulado, nUtilizado, nDisponible, tTipoCliente, tDistrito, tUsuario, fRegistro,testadofrecuente,lexcluyeproductos,lClienteCtaCte,nLinea,nConsumo,tTipoCtaCte,TsubTipoCtaCte, tcodigoCliente, tTipoIdentidad, tIdentidad, nLineaPorCobrar, nConsumoPorCobrar, lEmisionMonedaExtranjera, nDiasCredito) " & _
                                  "values ('" & NuevoCorrelativo & "', " & _
                                         " '" & cboTarjeta.BoundText & "', " & _
                                         " '" & txtTarjeta.Text & "', " & _
                                         " '" & txtApellido.Text & "', " & _
                                         " '" & txtNombre.Text & "', " & _
                                         " '" & txtDireccion.Text & "', " & _
                                         " '" & txtTelefono.Text & "', " & _
                                         " '" & txtReferencia.Text & "', " & _
                                         " '" & cboZona.BoundText & "', " & _
                                         " " & nDescuento & ", " & _
                                         " " & IIf(IsNull(dtpFecha.Value), "Null", "'" & Format(dtpFecha.Value, "yyyy/MM/dd") & "'") & ", " & _
                                         " '" & txtEmail.Text & "', " & _
                                         " '" & txtObservacion.Text & "', " & _
                                                chkActivo.Value & ", " & _
                                                chkPunto.Value & ", 0, 0, 0, " & _
                                         " '" & cboTipoCliente.BoundText & "', '" & cboDistrito.BoundText & "', " & _
                                           " '" & sUsuario & "', " & _
                                           " getdate() , '" & cboEstadoCliente.BoundText & "'," & chkExcluyeProductos.Value & "," & chkCuentaCorriente.Value & "," & nLinea & ",0,'" & cboTipoCtaCte.BoundText & "','" & cboSubTipoCtaCte.BoundText & "','" & sClienteFactura & "','" & cboTipoIdentidad.BoundText & "','" & txtIdentidad.Text & "'," & nLineaXCobrar & ",0, " & chkEmisionMonedaExtranjera.Value & ", " & nDiasCredito & " )"
                                           
                                Conex.Execute Isql
                            End If
                    End If ' TERMINAR LA CONDICION SI ESTA PRENDIDO EL CENTRAL INFOREST
                  
                
                   
                   '====================================
                   
                   
                   
                  'Cambiar el SQL
                  Isql = "insert into TDELIVERY( " & _
                          "tCodigodelivery,  tCodigoTarjeta, tNumeroTarjeta, tApellido, tNombre, tDireccion, tTelefono, tReferencia, tZona, nDescuento, fNacimiento, tEMail, tObservacion, lActivo, lPuntos, nAcumulado, nUtilizado, nDisponible, tTipoCliente, tDistrito, tUsuario, fRegistro,testadofrecuente,lexcluyeproductos,lClienteCtaCte,nLinea,nConsumo,tTipoCtaCte,TsubTipoCtaCte, tcodigoCliente, tTipoIdentidad, tIdentidad, nLineaPorCobrar, nConsumoPorCobrar, lEmisionMonedaExtranjera, nDiasCredito, lRangoHorarios, tHoraInicio, tHoraFin, tCanal1, tCanal2, tCanal3, tCanal4, tCanal5,lSoloxDia  ) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & cboTarjeta.BoundText & "', " & _
                                " '" & txtTarjeta.Text & "', " & _
                                " '" & txtApellido.Text & "', " & _
                                " '" & txtNombre.Text & "', " & _
                                " '" & txtDireccion.Text & "', " & _
                                " '" & txtTelefono.Text & "', " & _
                                " '" & txtReferencia.Text & "', " & _
                                " '" & cboZona.BoundText & "', " & _
                                " " & nDescuento & ", " & _
                                " " & IIf(IsNull(dtpFecha.Value), "Null", "'" & Format(dtpFecha.Value, "yyyy/MM/dd") & "'") & ", " & _
                                " '" & txtEmail.Text & "', " & _
                                " '" & txtObservacion.Text & "', " & _
                                       chkActivo.Value & ", " & _
                                       chkPunto.Value & ", 0, 0, 0, " & _
                                " '" & cboTipoCliente.BoundText & "', '" & cboDistrito.BoundText & "', " & _
                                  " '" & sUsuario & "', " & _
                                  " getdate() , '" & cboEstadoCliente.BoundText & "'," & chkExcluyeProductos.Value & "," & chkCuentaCorriente.Value & "," & nLinea & ",0,'" & cboTipoCtaCte.BoundText & "','" & cboSubTipoCtaCte.BoundText & "','" & sClienteFactura & "','" & cboTipoIdentidad.BoundText & "','" & txtIdentidad.Text & "'," & nLineaXCobrar & ",0, " & chkEmisionMonedaExtranjera.Value & ", " & nDiasCredito & " , " & lRangoHorarios.Value & ",'" & Format(tHoraInicio.Value, "HH:nn") & "','" & Format(tHoraFin.Value, "HH:nn") & "'," & tCanal1.Value & "," & tCanal2.Value & "," & tCanal3.Value & "," & tCanal4.Value & "," & tCanal5.Value & ", " & lSoloxDia.Value & " )"
                                  
                  Cn.Execute Isql
                                  
                  frmClienteDelivery.RsCabecera.Sort = "Codigo ASC"
                  frmClienteDelivery.RsCabecera.Requery
                  frmClienteDelivery.RsCabecera.MoveLast
                                                      
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmClienteDelivery.RsCabecera.RecordCount = 0, 0, frmClienteDelivery.RsCabecera.AbsolutePosition) & " de " & frmClienteDelivery.RsCabecera.RecordCount
               Else
                If cboTipoCliente.BoundText = "" Then MsgBox ("!No ha seleccionado Tipo de Cliente¡"): cboTipoCliente.SetFocus: Exit Sub
                  'Cambiar el SQL
                  If tValorDelivery = "1" Then
                        If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tCodigoDelivery<>'" & txtCodigo.Text & "' and tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                           MsgBox "Teléfono Existente", vbExclamation, sMensaje
                           Exit Sub
                        End If
                  End If
                  
                  Isql = "update TDELIVERY set " & _
                         "tApellido ='" & txtApellido.Text & "', " & _
                         "tNombre ='" & txtNombre.Text & "', " & _
                         "tZona ='" & cboZona.BoundText & "', " & _
                         "tReferencia ='" & txtReferencia.Text & "', " & _
                         "tDireccion ='" & txtDireccion.Text & "', " & _
                         "tCodigoTarjeta ='" & cboTarjeta.BoundText & "', " & _
                         "tNumeroTarjeta ='" & txtTarjeta.Text & "', " & _
                         "tTelefono ='" & txtTelefono.Text & "', " & _
                         "tObservacion ='" & txtObservacion.Text & "', " & _
                         "tEmail ='" & txtEmail.Text & "', " & _
                         "tEstadofrecuente ='" & cboEstadoCliente.BoundText & "', " & _
                         "nDescuento =" & nDescuento & ", " & _
                         "fNacimiento= " & IIf(IsNull(dtpFecha.Value), "null", "'" & Format(dtpFecha.Value, "yyyy/MM/dd") & "'") & ", " & _
                         "tTipoCliente='" & cboTipoCliente.BoundText & "', tDistrito='" & cboDistrito.BoundText & "', " & _
                         "lPuntos =" & chkPunto.Value & ", " & _
                         "tcodigocliente='" & sClienteFactura & "', " & _
                         "lActivo =" & chkActivo.Value & ", " & _
                         "lExcluyeProductos =" & chkExcluyeProductos.Value & ", " & _
                         "lClienteCtaCte =" & chkCuentaCorriente.Value & ", nLineaPorCobrar=" & nLineaXCobrar & ", " & _
                         "nLinea=" & nLinea & ", tTipoCtaCte='" & cboTipoCtaCte.BoundText & "',tSubTipoCtaCte='" & cboSubTipoCtaCte.BoundText & "', " & _
                         "fModificacion = getdate(), tUsuario='" & sUsuario & "', tTipoIdentidad= '" & cboTipoIdentidad.BoundText & "', tIdentidad='" & txtIdentidad.Text & "', " & _
                         "lEmisionMonedaExtranjera = " & chkEmisionMonedaExtranjera.Value & ", nDiasCredito = " & nDiasCredito & " " & _
                         ", lRangoHorarios =" & lRangoHorarios.Value & ", thorainicio ='" & Format(tHoraInicio.Value, "HH:nn") & "', thoraFin='" & Format(tHoraFin.Value, "HH:nn") & "' , tCanal1=" & tCanal1.Value & ", tCanal2=" & tCanal2.Value & ", tCanal3=" & tCanal3.Value & ", tCanal4=" & tCanal4.Value & ", tCanal5=" & tCanal5.Value & ",lSoloxdia =" & lSoloxDia.Value & " " & _
                         " where tCodigoDelivery = '" & txtCodigo & "'"
                   Cn.Execute Isql
                         
                        
                   nPos = frmClienteDelivery.RsCabecera.Bookmark
                   frmClienteDelivery.RsCabecera.Requery
                   If frmClienteDelivery.RsCabecera.RecordCount = 0 Then
                      frmClienteDelivery.RsCabecera.Filter = adFilterNone
                   End If
                   frmClienteDelivery.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                   
                  Cn.Execute "delete from tclienteproducto where tCodigoDelivery='" & txtCodigo & "'"
                  Cn.Execute "insert into tclienteproducto (tCodigoDelivery, tCodigoProducto, nPRECIO,lPermiteDescuentos,tUsuario, Fregistro) " & _
                             "select '" & txtCodigo.Text & "', tCodigoProducto, nPRECIO,LpERMITEDESCUENTO, '" & sUsuario & "', getdate() from " & sTemporal
                  RsProducto.Requery
                  
                  'deliverycliente
               If sClienteFactura <> "" Then
                   Cn.Execute "usp_Inforest_ValidaDeliveryCliente '" & txtCodigo.Text & "','" & sClienteFactura & "'"
               
               End If
               'deliverycliente
                  
                    Cn.Execute "delete from " & sTablaParientes
                    Cn.Execute "insert into " & sTablaParientes & " (Codigo,Pariente,Conyugue, Hijo) SELECT tcodigoPariente, pariente, lconyugue,lhijo from vPariente where tcodigodelivery='" & txtCodigo.Text & "' order by 2 asc"
                    rsParientes.Requery
       
                  
                  txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                   
          Case Is = 2 ' Eliminar
               If frmClienteDelivery.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Cliente " & txtApellido & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               If Calcular("select count(tClienteDelivery) as Codigo from MPEDIDO where tClienteDelivery='" & txtCodigo & "'", Cn) > 0 Or Calcular("select count(tClienteCtaCte) as Codigo from MPEDIDO where tClienteCtaCte='" & txtCodigo & "'", Cn) > 0 Then
                  MsgBox "Cliente con movimientos, no se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               Else
'                  'ClubNacional
'                  If (IIf(IsNull(lClub), False, lClub)) Then
'                        If Calcular("select count(tdocumento) as Codigo from MDOCUMENTO where tClientePago='" & txtCodigo & "'", Cn) > 0 Then
'                            MsgBox "Cliente con movimientos, no se puede eliminar", vbExclamation, sMensaje
'                            Exit Sub
'                        Else
'                            Cn.Execute "delete from TCOMPANIA where tCodigoCliente = '" & txtCodigo & "'"
'                        End If
'                  End If
               
                  Cn.Execute "delete from TDELIVERY where tCodigoDelivery = '" & txtCodigo & "'"
                  frmClienteDelivery.RsCabecera.Requery
               End If
               If frmClienteDelivery.RsCabecera.RecordCount <> 0 Then
                  frmClienteDelivery.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmClienteDelivery.RsCabecera.RecordCount = 0, 0, frmClienteDelivery.RsCabecera.AbsolutePosition) & " de " & frmClienteDelivery.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
                           
          Case Is = 4 ' tiendas
               If txtCodigo.Text = "" Then
                  MsgBox "Agregue primero al Cliente", vbCritical + vbOKOnly, sMensaje
                  Exit Sub
               End If
               sCodigo = txtCodigo.Text
               sDescrip = txtNombre.Text '& " " & txtNombre.Text ' txtDetallado.Text
               sTemp = txtApellido.Text '& " " & txtNombre.Text ' txtResumido.Text
               frmTienda.Show vbModal
          
                
          Case Is = 6
                If txtCodigo.Text <> "" Then
                
                    If sCodigo <> "" And Me.txtProducto.Caption <> "" And val(txtPrecio.Text) > 0 Then
                             If Calcular("select count(*) as codigo from " & sTemporal & " where tcodigoproducto='" & sCodigo & "'", Cn) = 0 Then
                                     Cn.Execute "insert into " & sTemporal & " (tcodigoDelivery,tCodigoProducto,nPrecio,lPermiteDescuento) values ('" & Me.txtCodigo.Text & "','" & sCodigo & "'," & val(Me.txtPrecio.Text) & "," & IIf(Me.chkExcluyente.Value, 1, 0) & ")"
                                     
                                     RsProducto.Requery
                                     
                                     txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                                     
                             Else
                                     MsgBox "Ya existe el Producto seleccionado", vbCritical, sMensaje
                             
                             End If
                             txtProducto.Caption = ""
                             txtPrecio.Text = ""
                             chkExcluyente.Value = 0
                             CmdBuscaProducto.SetFocus
                     End If
                Else
                MsgBox "Debe Grabar el Cliente para poder Asociarle Productos"
                
                End If
          Case Is = 9
              If RsProducto.RecordCount = 0 Then
                  Exit Sub
              End If
              Cn.Execute "delete " & sTemporal & " where tCodigoProducto ='" & RsProducto!tCodigoProducto & "'"
              RsProducto.Requery
              txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
   End Select

End Sub

Private Sub CmdBuscaProducto_Click()
        If txtCodigo.Text <> "" Then
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lActivo=1 and Codigo not in (select tCodigoProducto from " & sTemporal & " ) order by Descripcion"
               frmBusca.nPredeterm = 3
               Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                                      "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                                      "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                                      "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
               frmBusca.Show vbModal
               If Not wEnter Then
                    Me.txtProducto.Caption = ""
                  Exit Sub
               End If
               Me.txtProducto.Caption = sDescrip
               Me.txtPrecio.SetFocus
        Else
          MsgBox "Debe Grabar el Cliente para poder Asociarle Productos"
                
        End If
End Sub




 

 



 

Private Sub Form_Load()
    If VerConexionInfoRestCentral = "ON" Then
        BtnBuscarCentral.Visible = True
    Else
        BtnBuscarCentral.Visible = False
    End If
 


    If Sw = True Then
        
    
        Isql = "select  tDetallado , tCodigoCanalVenta from TCANALVENTA order by tCodigoCanalVenta asc"
        Set RsxCanalventa = Lib.OpenRecordset(Isql, Cn)
        
        RsxCanalventa.MoveFirst
        Do Until RsxCanalventa.EOF
             If RsxCanalventa!tCodigoCanalVenta = "01" Then
                tCanal1.Caption = RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "02" Then
                tCanal2.Caption = RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "03" Then
                tCanal3.Caption = RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "04" Then
                tCanal4.Caption = RsxCanalventa!tDetallado
             End If
             If RsxCanalventa!tCodigoCanalVenta = "05" Then
                tCanal5.Caption = RsxCanalventa!tDetallado
             End If
             
            RsxCanalventa.MoveNext
        Loop
        
        
       
        
        
        
        
'        tCanal2.Caption = "Aplica Canal - " & RsxCanalventa.Fields(1).Value
'        tCanal3.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado(2)
'        tCanal4.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado(3)
'        tCanal5.Caption = "Aplica Canal - " & RsxCanalventa!tDetallado(4)
    End If



    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Clientes Frecuentes "
    'fraDetalle.Caption = Me.
    
    'Llena todos los Combos
    LlenaCombos
        
    sTemporal = dbTemporal(sCaja, 4, "tcodigoDelivery", "nvarchar(9)", "tCodigoProducto", "nVarChar(7)", "nPrecio", "float", "lPermiteDescuento ", "bit")
    Isql = " SELECT     " & sTemporal & ".tcodigodelivery, " & sTemporal & ".tcodigoproducto, dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, dbo.vProducto.tResumido AS Producto, " & sTemporal & ".nPrecio, " & sTemporal & ".lPermiteDescuento " & _
           " FROM         dbo.vProducto INNER JOIN " & sTemporal & "  ON dbo.vProducto.codigo = " & sTemporal & ".tCodigoProducto  " & _
           " ORDER BY dbo.vProducto.Grupo, dbo.vProducto.SubGrupo, Producto"
           
    Set RsProducto = Lib.OpenRecordset(Isql, Cn)
    
    sTablaParientes = dbTemporal(sCaja, 4, "Codigo", "nvarchar(7)", "Pariente", "nVarChar(300)", "Conyugue", "bit", "Hijo", "bit")
    Isql = "select codigo as codigo, pariente as Pariente, Conyugue, Hijo from " & sTablaParientes & " order by pariente "
    Set rsParientes = Lib.OpenRecordset(Isql, Cn)
    
    ConfGrilla 4, grdParientes, "Codigo", 2, "Codigo", 1800, 2, 0, "", _
                             "Pariente", 2, "Pariente", 3800, 0, 0, "", _
                              "Conyugue", 2, "Conyugue", 1300, 2, 4, "", _
                              "Hijo", 2, "Hijo", 1300, 2, 4, ""
                            
                            
    Set grdParientes.DataSource = rsParientes
    
    ConfGrilla 6, grdGrilla, "Codigo", 2, "tCodigoProducto", 800, 2, 0, "", _
                             "Grupo", 2, "Grupo", 1100, 0, 0, "", _
                             "SubGrupo", 2, "SubGrupo", 1250, 0, 0, "", _
                             "Producto", 2, "Producto", 3200, 0, 0, "", _
                             "Precio", 2, "nPrecio", 850, 1, 0, "###,##0.00", _
                             "Descuento", 2, "lPermiteDescuento", 1000, 2, 4, ""

    Set grdGrilla.DataSource = RsProducto
      
 
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       dtpFecha.Value = Null
       txtDescuento.Text = "0.00"
       cboEstadoCliente.BoundText = "01"
       chkEmisionMonedaExtranjera.Value = 0
    Else
       'Cambiar el Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
        
    txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    cmdTexto.Caption = "Registro " & IIf(frmClienteDelivery.RsCabecera.RecordCount = 0, 0, frmClienteDelivery.RsCabecera.AbsolutePosition) & " de " & frmClienteDelivery.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmClienteDeliveryDetalle = Nothing
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

 Private Sub lRangoHorarios_Click()
If lRangoHorarios Then
    tHoraInicio.Enabled = True
    tHoraFin.Enabled = True
Else
    tHoraInicio.Enabled = False
    tHoraFin.Enabled = False
End If


End Sub

 
 
Private Sub tHoraFin_Change()
 If tHoraInicio.Value > tHoraFin.Value Then
      tHoraFin.Value = tHoraFin.Value
      MsgBox "Error en Horas", vbExclamation, sMensaje
   End If
End Sub

Private Sub txtApellido_LostFocus()
   Call ValidaStr(txtApellido)
End Sub

Private Sub txtCuentaCobrar_LostFocus()
   nLineaXCobrar = val(txtCuentaCobrar.Text)
    txtCuentaCobrar.Text = Format(nLineaXCobrar, "######0.00")
End Sub

Private Sub txtDescuento_GotFocus()
   txtDescuento.SelLength = Len(Trim(txtDescuento.Text))
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
   Numerico KeyAscii, txtDescuento
End Sub

Private Sub txtDescuento_LostFocus()
    nDescuento = val(txtDescuento.Text)
    txtDescuento.Text = Format(nDescuento, "##0.00")
End Sub

Private Sub txtDiasCredito_LostFocus()
    nDiasCredito = val(txtDiasCredito.Text)
    txtDiasCredito.Text = Format(nDiasCredito, "##0")
End Sub

Private Sub txtDireccion_LostFocus()
   Call ValidaStr(txtDireccion)
End Sub

Private Sub txtEmail_LostFocus()
   Call ValidaStr(txtEmail)
End Sub

Private Sub txtNombre_LostFocus()
   Call ValidaStr(txtNombre)
End Sub

Private Sub txtObservacion_LostFocus()
   Call ValidaStr(txtObservacion)
End Sub

Private Sub txtReferencia_LostFocus()
   Call ValidaStr(txtReferencia)
End Sub

Private Sub txtTelefono_LostFocus()
   Call ValidaStr(txtTelefono)
End Sub

Private Sub txtLinea_LostFocus()
    nLinea = val(txtLinea.Text)
    txtLinea.Text = Format(nLinea, "######0.00")
End Sub

Private Sub cboTipoCtaCte_Click(Area As Integer)
    cboSubTipoCtaCte.BoundText = ""
    With cboSubTipoCtaCte
         Isql = "Select * from vSubtipoctacte where tTipoCtaCte = '" & cboTipoCtaCte.BoundText & "' order by Descripcion "
         Set RsSubTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipoCtaCte
    End With
End Sub

Private Sub txtPrecio_GotFocus()
   Enfoque txtPrecio
End Sub

Private Sub txtPrecio_LostFocus()
   nPrecio = val(txtPrecio.Text)
   txtPrecio.Text = Format(nPrecio, "###0.00")
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtPrecio
End Sub

