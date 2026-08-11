VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReservaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11205
   Icon            =   "frmReservaDetalleAdicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11205
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
      Height          =   5220
      Left            =   2760
      TabIndex        =   23
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton Command3 
         Caption         =   "PAX:"
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Mesa:"
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "obs :"
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
         Left            =   240
         MaskColor       =   &H8000000A&
         TabIndex        =   40
         Top             =   3480
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Anticipos"
         ForeColor       =   &H00FF0000&
         Height          =   4215
         Left            =   4200
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Agregar Anticipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   2760
            Picture         =   "frmReservaDetalleAdicion.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   3480
            Width           =   1170
         End
         Begin TrueOleDBGrid80.TDBGrid grdPedido 
            Height          =   3195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5636
            _LayoutType     =   4
            _RowHeight      =   25
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
            Caption         =   "Anticipos"
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
         Begin VB.Label lblTotal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   3720
            Width           =   2535
         End
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
         Left            =   7560
         Picture         =   "frmReservaDetalleAdicion.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdBuscaMesero 
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
         Left            =   3600
         Picture         =   "frmReservaDetalleAdicion.frx":0A6E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtMesero 
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
         Height          =   345
         Left            =   960
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtMesa 
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
         Height          =   345
         Left            =   960
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   2040
         Width           =   1530
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   345
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   6600
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
         Height          =   345
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   2130
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
         Height          =   885
         Left            =   960
         MaxLength       =   80
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmReservaDetalleAdicion.frx":0B70
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox txtRegistro 
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
         Height          =   345
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   24
         Top             =   240
         Width           =   2490
      End
      Begin VB.TextBox txtTelefono 
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
         Height          =   345
         Left            =   960
         MaxLength       =   15
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   2370
      End
      Begin VB.TextBox txtPax 
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
         Height          =   345
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   4
         Text            =   " "
         Top             =   2040
         Width           =   690
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   345
         Left            =   2520
         TabIndex        =   3
         Top             =   1560
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   55508995
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   345
         Left            =   960
         TabIndex        =   2
         Top             =   1560
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55508995
         CurrentDate     =   37795
      End
      Begin MSDataListLib.DataCombo cboMotivo 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   2520
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mesero :"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora :"
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mesa :"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro :"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   33
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   800
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   800
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   800
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   800
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PAX :"
         Height          =   195
         Index           =   7
         Left            =   2640
         TabIndex        =   27
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Atendido"
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
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11145
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5265
      Width           =   11205
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Atender Reserva"
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
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
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
         Left            =   4620
         Picture         =   "frmReservaDetalleAdicion.frx":0B72
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
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
         Left            =   8130
         Picture         =   "frmReservaDetalleAdicion.frx":10A4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
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
         Left            =   9300
         Picture         =   "frmReservaDetalleAdicion.frx":11A6
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   5790
         Picture         =   "frmReservaDetalleAdicion.frx":1298
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4500
         TabIndex        =   15
         Top             =   60
         Width           =   4560
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmReservaDetalleAdicion.frx":17CA
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmReservaDetalleAdicion.frx":1D0C
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmReservaDetalleAdicion.frx":224E
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4005
            Picture         =   "frmReservaDetalleAdicion.frx":2790
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3525
            Picture         =   "frmReservaDetalleAdicion.frx":2CD2
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3045
            Picture         =   "frmReservaDetalleAdicion.frx":3214
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1440
            TabIndex        =   22
            Top             =   180
            Width           =   1545
         End
      End
   End
   Begin VB.Image Image 
      Height          =   5205
      Left            =   15
      Picture         =   "frmReservaDetalleAdicion.frx":3756
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2655
   End
End
Attribute VB_Name = "frmReservaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim xEstado As String
Dim sCodigoCliente As String
Dim sCodigoMesero As String
Dim RsMotivo As Recordset
Dim rsDocAnticipo As Recordset

Sub Asignar()
    With frmReserva.RsCabecera
        If .RecordCount <> 0 Then
            'Cuadro de Texto
            
            
            txtCodigo = IIf(IsNull(!TRESERVA), "", !TRESERVA)
            txtRegistro = IIf(IsNull(!fRegistro), FechaServidor(), !fRegistro)
            
            sCodigoCliente = IIf(IsNull(!tCodigoDelivery), "", !tCodigoDelivery)
            sCliente = sCodigoCliente
            If sCodigoCliente = "" Then
                txtCliente = IIf(IsNull(!tApellido), "", !tApellido) + IIf(IsNull(!tNombre), "", !tNombre)
            Else
                txtCliente = Calcular("select tApellido + ' ' + tNombre as codigo from vDelivery where Codigo='" & sCodigoCliente & "'", Cn)
            End If
        
            'txtApellido = IIf(IsNull(!tApellido), "", !tApellido)
            cboMotivo.BoundText = IIf(IsNull(!tMotivo), "01", !tMotivo)
            
            sCodigoMesero = IIf(IsNull(!tMesero), "", !tMesero)
            If sCodigoMesero = "" Then
                txtMesero = ""
            Else
                txtMesero = Calcular("select Descripcion as codigo from vMozo where left(Codigo,1) <> '*' and Codigo='" & sCodigoMesero & "'", Cn)
            End If
            
            dtpFecha.value = IIf(IsNull(!fFecha), FechaServidor(), !fFecha)
            dtpHora.value = IIf(IsNull(!fHora), Time, !fHora)
            txtTelefono = IIf(IsNull(!tTelefono), "", !tTelefono)
            txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
            txtPax = IIf(IsNull(!nPax), 0, !nPax)
            txtMesa = IIf(IsNull(!tMesa), 0, !tMesa)
            lblEstado.Caption = IIf(IsNull(!EstadoReserva), "", !EstadoReserva)
            xEstado = IIf(IsNull(!tEstadoReserva), "", !tEstadoReserva)
            If xEstado = "02" Then
                fraDetalle.Enabled = False
            ElseIf xEstado = "01" Then
                fraDetalle.Enabled = True
            End If
            
                Set rsDocAnticipo = Nothing
                
                Isql = "SELECT tdocumento , nventa  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "'"
                Set rsDocAnticipo = Lib.OpenRecordset(Isql, Cn)
                                                  
                Set grdPedido.DataSource = rsDocAnticipo
            
                If rsDocAnticipo.RecordCount > 0 Then
                   lblTotal.Caption = "Total Anticipo: " & Format(Calcular("SELECT sum(nventa) as codigo  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "'", Cn), "#,###,##0.00")
                Else
                    lblTotal.Caption = ""
                End If
               
                If Not lActivaAnticipo Or Trim(txtCodigo.Text) = "" Then
                   Frame1.Enabled = False
                Else
                   Frame1.Enabled = True
                End If
            
            
            cmdTexto.Caption = "Registro " & frmReserva.RsCabecera.AbsolutePosition & " de " & frmReserva.RsCabecera.RecordCount
        End If
    End With
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "select Codigo, tApellido + ' ' + tNombre as Descripcion,tTelefono as Telefono  from vDelivery"
   frmClienteDeliveryFrecuente.Show vbModal
   
   If Not wEnter Then
      sCodigo = ""
      sDescrip = ""
      txtCliente.Text = ""
      txtTelefono.Text = ""
      Exit Sub
   Else
        sCodigoCliente = sCodigo
        txtCliente.Text = sDescrip
        txtTelefono.Text = sTelefonoReserva
   End If

   
End Sub

Private Sub cmdBuscaMesero_Click()
    Dim xCriterio As String
    Isql = "Select Codigo, Descripcion, tResumido as Usuario from vMozo where lActivo=1 and left(Codigo,1) <> '*'"
    
    frmBusca.cboCriterio.Enabled = True
    frmBusca.nPredeterm = 1
    Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Descripcion", 2, "Descripcion", 3500, 0, 0, "", _
                                          "Usuario", 2, "Usuario", 2000, 0, 0, "")
    frmBusca.Show vbModal
    If Not wEnter Then
      Exit Sub
    End If
    sCodigoMesero = sCodigo
    txtMesero.Text = sDescrip
End Sub

Sub LlenaCombos()
    With cboMotivo
         Isql = "Select tCodigo as Codigo, tDetallado as Descripcion from TTABLA where TTABLA='MOTIVORESERVA' and lActivo=1"
         Set RsMotivo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotivo
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmReserva.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmReserva.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmReserva.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmReserva.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmReserva.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmReserva.grdGrilla
    End Select
   Asignar

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               lblEstado.Caption = ""
               Blanquear Me
               lblTotal.Caption = ""
                Me.grdPedido.DataSource = Nothing
                
               'Cambia el Nombre del Primer Text
               'txtApellido.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Integer
               
               
                If lReservaPedido Then
                    Dim EstadoReserva2  As String
                    EstadoReserva2 = Calcular("select testadoreserva as codigo from treserva where tReserva='" & Me.txtCodigo.Text & "'", Cn)
                    If EstadoReserva2 = "02" Then
                        MsgBox "No se puede modificar la reserva: " & txtCodigo.Text & vbNewLine & ", ya ha sido Atendido!!!", vbExclamation, sMensaje
                        Exit Sub
                    End If
                    If EstadoReserva2 = "03" Then
                        MsgBox "Reserva: " & txtCodigo.Text & vbNewLine & ", ya ha sido Eliminada!!!", vbExclamation, sMensaje
                        Exit Sub
                    End If
                End If
               
               If txtCliente.Text = "" Then MsgBox "Ingrese el Cliente", vbExclamation, sMensaje: Exit Sub
               If txtPax.Text = "" Then MsgBox "Ingrese el N° de Pax", vbExclamation, sMensaje: txtPax.SetFocus: Exit Sub
               nPos = frmReserva.RsCabecera.AbsolutePosition
               'Chequea Datos
               If Sw Then
                
                  sCodigo = Calcular("select Count(*) as Codigo from tRESERVA where tReserva = '" & txtCodigo.Text & "'", Cn)
                  
                  If sCodigo > 0 Then
                    MsgBox "Reserva: " & txtCodigo.Text & vbNewLine & ", ya ha sido registrada!!!", vbExclamation, sMensaje
                        Exit Sub
                  End If
               
                  'Obtiene el Numero de Reserva
                  
                  sCodigo = Calcular("select max(tReserva) as Codigo from tRESERVA where substring(tReserva,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(sCodigo) Or Mid(sCodigo, 1, 2) <> Mid(Trim(str(Year(FechaServidor()))), 3, 2) Then
                      sCodigo = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & "00000001"
                  Else
                      sCodigo = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & Lib.Correlativo(Mid(sCodigo, 3, 8), 8)
                  End If
                  Sw = False
                   
                  txtCodigo.Text = sCodigo
                  'Cambiar el SQL
                  Isql = "insert into TRESERVA( " & _
                         "tReserva, fRegistro, tCodigoDelivery, tTelefono, tMotivo, tMesero, fFecha, fHora, nPax, tEstadoReserva, tObservacion, tPrioridad, tPrioridad2, tMesa,fDiaContable) " & _
                         "values ( '" & txtCodigo.Text & "', " & _
                                 " getdate() , " & _
                                 " '" & sCodigoCliente & "', " & _
                                 " '" & txtTelefono.Text & "', " & _
                                 " '" & cboMotivo.BoundText & "', " & _
                                 " '" & sCodigoMesero & "', " & _
                                 " '" & Format(dtpFecha.value, "yyyy/MM/dd") & "', " & _
                                 "'" & Format(dtpHora.value, "hh:nn") & "', " & _
                                        Val(txtPax.Text) & ", " & _
                                 " '01', " & _
                                 " '" & Left(txtObservacion.Text, 200) & "'," & _
                                 " '1'," & _
                                 " '0'," & _
                                 " '" & txtMesa & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
           
                      Cn.Execute Isql
                      frmReserva.RsCabecera.Requery
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      'Asignar
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmReserva.RsCabecera.RecordCount = 0, 0, frmReserva.RsCabecera.AbsolutePosition) & " de " & frmReserva.RsCabecera.RecordCount
              Else
                 'Cambiar el SQL
                 nPos = frmReserva.RsCabecera.AbsolutePosition
                 Isql = "update TRESERVA set " & _
                        "tCodigoDelivery ='" & sCodigoCliente & "', " & _
                        "tTelefono ='" & txtTelefono.Text & "', " & _
                        "tMotivo ='" & cboMotivo.BoundText & "', " & _
                        "tMesero ='" & sCodigoMesero & "', " & _
                        "fFecha ='" & Format(dtpFecha.value, "yyyy/MM/dd") & "', " & _
                        "fHora ='" & Format(dtpHora.value, "hh:nn") & "', " & _
                        "nPax =" & Val(txtPax.Text) & ", " & _
                        "tObservacion ='" & Left(txtObservacion.Text, 200) & "', " & _
                        "tMesa ='" & txtMesa.Text & "' " & _
                        " where tReserva = '" & txtCodigo.Text & "'"
                                              
                  Cn.Execute Isql
                  nPos = frmReserva.RsCabecera.AbsolutePosition
                  frmReserva.RsCabecera.Requery
              End If
              
                '******************************************
                Set rsDocAnticipo = Nothing
                
                Isql = "SELECT tdocumento , nventa  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "' and tEstadoDocumento<>'04'"
                Set rsDocAnticipo = Lib.OpenRecordset(Isql, Cn)
                                                  
                Set grdPedido.DataSource = rsDocAnticipo
            
                If rsDocAnticipo.RecordCount > 0 Then
                   lblTotal.Caption = "Total Anticipo: " & Format(Calcular("SELECT sum(nventa) as codigo  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "' and tEstadoDocumento<>'04' ", Cn), "#,###,##0.00")
                End If
               
                If Not lActivaAnticipo Or Trim(txtCodigo.Text) = "" Then
                   Frame1.Enabled = False
                Else
                   Frame1.Enabled = True
                End If
                '*******************************************
                Unload Me
             'frmReserva.RsCabecera.AbsolutePosition = IIf(nPos <= 0, 1, nPos)
          Case Is = 2 ' Eliminar
               If frmReserva.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               If xEstado = "03" Then
                    MsgBox "Reserva: " & txtCodigo.Text & vbNewLine & ", ya ha sido Eliminada!!!", vbExclamation, sMensaje
                    Exit Sub
                End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Anular la Reserva " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "update TRESERVA set tEstadoReserva ='03' where tReserva = '" & txtCodigo.Text & "'"
               frmReserva.RsCabecera.Requery
               If Not frmReserva.RsCabecera.EOF Then
                  frmReserva.RsCabecera.MoveFirst
                  Asignar
               Else
                  Unload Me
               End If
                                       
          Case Is = 3 ' Salir
               Unload Me
               
          Case Is = 4 ' Estado
               'Cambiar el SQL

               If lReservaPedido Then
                    Dim EstadoReserva, PedidoReserva  As String
                    EstadoReserva = Calcular("select testadoreserva as codigo from treserva where tReserva='" & Me.txtCodigo.Text & "'", Cn)
                    PedidoReserva = Calcular("select tCodigoPedido as codigo from mpedido where treservaInf='" & Me.txtCodigo.Text & "' and ffecha>= getdate()-90  and testadopedido<>'03'", Cn)
                    If EstadoReserva = "02" Then
                        MsgBox "La Reserva: " & txtCodigo.Text & vbNewLine & ", ya se ha atendido!!!", vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    If EstadoReserva = "03" Then
                        MsgBox "La reserva " & Me.txtCodigo.Text & " esta Anulada!!!", vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    If EstadoReserva = "01" Then
                      Dim oComando As clsComando
                      Set oComando = New clsComando
                        If Not oComando.CreateCmdSp("spIns_MPEDIDO_RESERVA", Cn) Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                    
                      oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                      oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                      oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
                      oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
                      oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(dtpFecha.value = "", Null, Format(dtpFecha.value, "dd/MM/yyyy HH:mm"))
                      oComando.CreateParameter "@codigoReserva", adVarChar, adParamInput, 10, txtCodigo.Text
                      oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, ""
                     If Not oComando.GetParamOK Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                    If Not oComando.ExecSP Then
                       Set oComando = Nothing
                       Exit Sub
                    Else
                       sPedido = oComando.GetParameterValue("@tPedido")
                    End If
                  
                  If sPedido = "" Then
                    MsgBox "Hubo un error en la generacion del pedido de la reserva: " & txtCodigo.Text & vbNewLine & "Favor de volver a intentarlo o consultar a su administrador!!!", vbExclamation, sMensaje
                    Exit Sub
                  End If
                  
                End If
                
                Select Case EstadoReserva
                  Case Is = "01"
                       Cn.Execute "update TRESERVA set tEstadoReserva ='02' where tReserva = '" & txtCodigo.Text & "'"
                  Case Is = "02"
                       'Cn.Execute "update TRESERVA set tEstadoReserva ='01' where tReserva = '" & txtCodigo.Text & "'"
                  Case Is = "03"
                       MsgBox "Reserva Anulada, no se puede cambiar el estado", vbExclamation, sMensaje
                       Exit Sub
                  End Select
               Else
                    Dim EstadoReser As String
                    EstadoReser = Calcular("select testadoreserva as codigo from treserva where tReserva='" & Me.txtCodigo.Text & "'", Cn)
                    
                    Select Case EstadoReser
                      Case Is = "01"
                           Cn.Execute "update TRESERVA set tEstadoReserva ='02' where tReserva = '" & txtCodigo.Text & "'"
                      Case Is = "02"
                            Dim PedidoReser As String
                            PedidoReser = Calcular("select tCodigoPedido as codigo from mpedido where treservaInf='" & Me.txtCodigo.Text & "' and ffecha>= getdate()-90  and testadopedido<>'03'", Cn)
                            If PedidoReser <> "" And PedidoReser <> "0" Then
                                MsgBox "La Reserva: " & txtCodigo.Text & vbNewLine & ", ya se ha atendido con el pedido: " & PedidoReser & "!!!", vbExclamation, sMensaje
                                Exit Sub
                            Else
                               Cn.Execute "update TRESERVA set tEstadoReserva ='01' where tReserva = '" & txtCodigo.Text & "'"
                            End If
                           
                      Case Is = "03"
                           MsgBox "Reserva Anulada, no se puede cambiar el estado", vbExclamation, sMensaje
                           Exit Sub
                      End Select
               End If
               
               nPos = frmReserva.RsCabecera.AbsolutePosition
               frmReserva.RsCabecera.Requery
               If Not frmReserva.RsCabecera.EOF Then
                  frmReserva.RsCabecera.MoveFirst
                  Asignar
               Else
                  Unload Me
               End If
               If frmReserva.RsCabecera.AbsolutePosition > 0 Then
                frmReserva.RsCabecera.AbsolutePosition = nPos
               End If
               MsgBox "Registro Modificado", vbInformation, sMensaje
               Unload Me
               
        Case Is = 5
'
'            If lCajaContingencia Then 'GCAA 22032021
'                MsgBox "No se puede agregar Anticipos, cuando la caja esta configurada como Caja Contingencia", vbExclamation, sMensaje
'                Exit Sub
'            End If
'
'
'            If txtCodigo.Text <> "" Then
'                sCliente = sCodigoCliente
'                Sw = True
'                frmReciboIngresoDetalle.TipoOper = "Anticipo"
'                frmReciboIngresoDetalle.Show vbModal
'                frmReciboIngresoDetalle.TipoOper = ""
'            Else
'                MsgBox "Es necesario generar una Reserva!!!"
'            End If
'
'            rsDocAnticipo.Requery
'
'            If rsDocAnticipo.RecordCount > 0 Then
'                lblTotal.Caption = "Total Anticipo: " & Format(Calcular("SELECT sum(nventa) as codigo  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "'", Cn), "#,###,##0.00")
'            End If
        
   End Select
End Sub

Private Sub Command1_Click()
    frmKeyBoard.txtResultado = txtObservacion.Text
    frmKeyBoard.Show vbModal
    txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
End Sub

Private Sub Command2_Click()
    frmKeyBoard.txtResultado = txtMesa.Text
    frmKeyBoard.Show vbModal
    txtMesa.Text = IIf(wEnter, sDescrip, txtMesa.Text)
End Sub

Private Sub Command3_Click()
    sTipo = "Numero"
    frmNumPad.Show vbModal
    txtPax.Text = IIf(wEnter = True, sDescrip, txtPax.Text)
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    sCodigoCliente = ""
    sCodigoMesero = ""
    
    LlenaCombos
    
    Label(10).Caption = Mesero
    
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Reservas "
    fraDetalle.Caption = Me.Caption
    dtpFecha.value = FechaServidor()
          
    If Sw = True Then
       ActivarBotones (False)
       lblEstado.Caption = ""
       Blanquear Me
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If

   Call ConfGrilla(2, grdPedido, "Documento", 2, "tdocumento", 1500, 1, 0, "", _
                                 "Total", 2, "nventa", 1200, 1, 0, "#,###,##0.00")
        Isql = "SELECT tdocumento , nventa  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "' and isnull(TRESERVA,'')<>''"
        Set rsDocAnticipo = Lib.OpenRecordset(Isql, Cn)
                                      
        Set grdPedido.DataSource = rsDocAnticipo
        If rsDocAnticipo.RecordCount > 0 Then
            lblTotal.Caption = "Total Anticipo: " & Format(Calcular("SELECT sum(nventa) as codigo  from MDOCUMENTO WHERE TRESERVA='" & txtCodigo.Text & "'", Cn), "#,###,##0.00")
        End If

    
    If Not lActivaAnticipo Or Trim(txtCodigo.Text) = "" Then
       Frame1.Enabled = False
    Else
       Frame1.Enabled = True
    End If
    
    
    
    
    cmdTexto.Caption = "Registro " & frmReserva.RsCabecera.AbsolutePosition & " de " & frmReserva.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReservaDetalle = Nothing
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


