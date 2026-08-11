VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReserva 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8595
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   11850
   Icon            =   "frmReservaAdicion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8566.058
   ScaleMode       =   0  'User
   ScaleWidth      =   11850
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   11790
      TabIndex        =   2
      Top             =   7170
      Width           =   11850
      Begin VB.CommandButton btnRevertirPedido 
         Caption         =   "Revertir Pedido"
         Height          =   555
         Left            =   7080
         TabIndex        =   25
         Top             =   90
         Width           =   1695
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
         Height          =   525
         Index           =   3
         Left            =   9120
         Picture         =   "frmReservaAdicion.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   10320
         Picture         =   "frmReservaAdicion.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   1245
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   15
         Top             =   30
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmReservaAdicion.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmReservaAdicion.frx":0B88
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmReservaAdicion.frx":10CA
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmReservaAdicion.frx":160C
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmReservaAdicion.frx":1B4E
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmReservaAdicion.frx":2090
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
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
            Height          =   315
            Left            =   1650
            TabIndex        =   22
            Top             =   150
            Width           =   2505
         End
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   0
         Left            =   5895
         Picture         =   "frmReservaAdicion.frx":25D2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agrega"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   1
         Left            =   6480
         Picture         =   "frmReservaAdicion.frx":2B04
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Modifica"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdEmite 
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
         Left            =   7065
         Picture         =   "frmReservaAdicion.frx":2C06
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Emite"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   6
         Left            =   8820
         Picture         =   "frmReservaAdicion.frx":3138
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   3840
         Picture         =   "frmReservaAdicion.frx":322A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filtrar"
         Top             =   720
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
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
         Left            =   165
         TabIndex        =   5
         Top             =   650
         Width           =   3660
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2205
            TabIndex        =   6
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   16384001
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   7
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   16384001
            CurrentDate     =   37539
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Del "
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
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " al "
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
            Height          =   195
            Index           =   1
            Left            =   1845
            TabIndex        =   8
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Enabled         =   0   'False
         Height          =   555
         Index           =   5
         Left            =   8235
         Picture         =   "frmReservaAdicion.frx":332C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Reportes"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7650
         Picture         =   "frmReservaAdicion.frx":385E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar HTML"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6870
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12118
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
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
         Splits(0).FilterBar=   -1  'True
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
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
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
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   225
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim rsReporte As Recordset
'Dim Reporte As New dsrReserva
Dim nColumna As Integer

Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Private Sub btnRevertirPedido_Click()  'Revertir reserva a estado generado
    If RsCabecera.RecordCount > 0 Then
        With frmReserva.RsCabecera
        
            Dim tcodigoReserva, tcodPedido, tcMesa As String
            Dim countDPedido, countMPedido As Integer
    
            tcodigoReserva = !TRESERVA
            tcodPedido = !PedidoInfoRest
            
            If UCase(!EstadoReserva) <> "ATENDIDO" Then
                MsgBox "Esta reserva no puede ser revertida; únicamente se pueden revertir aquellas reservas que se encuentren con el estado 'Atendido'.", vbExclamation, sMensaje
                Exit Sub
            End If
          
            countDPedido = Calcular("select count(*) as codigo from DPEDIDO where tCodigoPedido='" & tcodPedido & "' and (ISNULL(lImprime,0)<>0 or ISNULL(nEnvio,0)<>0)", Cn)

            If countDPedido > 0 Then
               MsgBox "Error: No es posible revertir la reserva de este pedido, existen productos enviados", vbExclamation, sMensaje
               Exit Sub
            End If

            countMPedido = Calcular("select Count(*) as codigo from MPEDIDO where tEstadoPedido in ('01','03') and fRegistro>CONVERT(date,GETDATE()) and tCodigoPedido='" & tcodPedido & "'", Cn)

            If countMPedido = 0 Then
               MsgBox "Error: No es posible revertir la reserva de este pedido.", vbExclamation, sMensaje
               Exit Sub
            End If

            If MsgBox("Seguro de revertir la Reserva " & tcodigoReserva & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               Exit Sub
            End If
            
            Cn.Execute "Delete from DPEDIDO where tCodigoPedido= '" & tcodPedido & "'"

            Cn.Execute "Update MPEDIDO Set tEstadoPedido='03', tObservacionAnulado='Se anuló por generar la reversión de la reserva.' Where tCodigoPedido='" & tcodPedido & "'"
            
            tcMesa = Calcular("select ISNULL(tMesa,'') as codigo from MPEDIDO where tEstadoPedido in ('01') and tCodigoPedido='" & tcodPedido & "'", Cn)
             
            If tcMesa <> "" Or tcMesa <> "000" Then
               Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa ='" & tcMesa & "'"
            End If
            
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & tcodPedido & "')"
            Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido='" & tcodPedido & "'"
   
   
            Cn.Execute "update TRESERVA_COVERMANAGER set PedidoInforest = '' where CodigoReservaInfoRest = '" & tcodigoReserva & "'"
            
            Cn.Execute "UPDATE TRESERVA SET tEstadoReserva = '01',fFechaModificacion = GETDATE() where tReserva = '" & tcodigoReserva & "'"
            
            frmCargoMozo.RsCabecera.Requery
            frmCargoMozo.LimpiaCabecera
            frmReserva.RsCabecera.Requery
            Unload Me
             
        End With
    Else
        MsgBox "No Existe Datos", vbExclamation, sMensaje
    End If
End Sub

Private Sub cmdEmite_Click()
    With grdGrilla.PrintInfo
        'Cabecera
        .SettingsOrientation = 2
        .PageHeaderFont.Italic = True
        .RepeatColumnHeaders = True
        .PreviewMaximize = True
        .SettingsMarginLeft = 400
        .SettingsMarginRight = 400
        .SettingsMarginTop = 400
        .SettingsMarginBottom = 400
        .VariableRowHeight = True
        .PageFooter = "\tPage: \p"
        .PreviewCaption = "Emite de " & Me.Caption
        .PrintPreview
    End With
End Sub

Private Sub cmdExporta_Click()
    cmdSave.Filter = "Hyper Text Markup Language|*.html"
    cmdSave.ShowSave
    If cmdSave.FileName = "" Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    grdGrilla.ExportToFile cmdSave.FileName, False, False
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_DblClick()
  On Error GoTo errHandler
'
    ' Declara las variables    Set cols = grdGrilla.Columns
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col ' Guarda la columna activa

'    Dim indiceFila As Long
'    indiceFila = grdGrilla.Row



    If c = 2 Then
'        If grdGrilla.Text <> "" Then
            ' Guarda el texto de la celda actual y muestra el formulario del teclado
'            frmKeyBoard.txtResultado = grdGrilla.Text ' Almacena el texto de la celda activa en el formulario
            If frmKeyBoard.Visible = False Then
                frmKeyBoard.Caption = "Digite el nombre del cliente"
                frmKeyBoard.Show vbModal

            End If


            ' Verifica si el valor de wEnter es verdadero, lo que indica que se presionó Enter
            If wEnter = True Then
                ' Asegúrate de que la fila esté seleccionada correctamente

                If sDescrip <> "" Then
                    grdGrilla.HoldFields
                    RsCabecera.Filter = "Cliente LIKE '%" & sDescrip & "%'"
                    grdGrilla.col = c
                    grdGrilla.EditActive = True

                    ' Verificar si el filtro se aplicó correctamente
                    If RsCabecera.EOF Then
                        For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                        MsgBox "No se encontró el cliente con el nombre: " & sDescrip, vbInformation, "Búsqueda"
                        Call cmdClearFilter_Click(cols)
                        RsCabecera.Filter = adFilterNone
                    Else

                        For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                        ' Si se encuentran registros, accede al primer registro
                       RsCabecera.MoveFirst

                       ' Cambia el valor en el Recordset (por ejemplo, en la columna "Cliente")
    '                   RsCabecera.Fields("Cliente").value = nuevoValor ' Sustituye "NuevoValor" con el valor que deseas asignar
                        RsCabecera.Update

                        grdGrilla.HoldFields    '
                        grdGrilla.col = c
                        grdGrilla.EditActive = True

                        grdGrilla.Refresh
    '
                    End If
                Else
                    For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                    grdGrilla.HoldFields
                    RsCabecera.Filter = getFilter(cols)
                    grdGrilla.col = c
                    grdGrilla.EditActive = True
                End If
            End If
'        Else
'            grdGrilla.HoldFields
'            RsCabecera.Filter = getFilter(cols)
'            grdGrilla.col = c
'            grdGrilla.EditActive = True
'        End If
    Else
        grdGrilla.HoldFields
        RsCabecera.Filter = getFilter(cols)
        grdGrilla.col = c
        grdGrilla.EditActive = True

    End If

    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje

    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    Else
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdGrilla_FilterChange()
'    On Error GoTo errHandler
'    Set cols = grdGrilla.Columns
'    Dim c As Integer
'    c = grdGrilla.col
'    grdGrilla.HoldFields
'    RsCabecera.Filter = getFilter(cols)
'    grdGrilla.col = c
'    grdGrilla.EditActive = True
'    Exit Sub
'
'errHandler:
'    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
'
'    Call cmdClearFilter_Click(cols)
'    RsCabecera.Filter = adFilterNone
 On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    ' Declara las variables    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col ' Guarda la columna activa
    
    If c = 2 Then
        If grdGrilla.Text <> "" Then
            ' Guarda el texto de la celda actual y muestra el formulario del teclado
            If frmKeyBoard.Visible = False Then
                frmKeyBoard.txtResultado = grdGrilla.Text ' Almacena el texto de la celda activa en el formulario
                frmKeyBoard.Caption = "Digite el nombre del cliente"
                frmKeyBoard.Show vbModal
            
            End If
            
            
            ' Verifica si el valor de wEnter es verdadero, lo que indica que se presionó Enter
            If wEnter = True Then
                ' Asegúrate de que la fila esté seleccionada correctamente
                
                If sDescrip <> "" Then
                    grdGrilla.HoldFields
                    RsCabecera.Filter = "Cliente LIKE '%" & sDescrip & "%'"
                    grdGrilla.col = c
                    grdGrilla.EditActive = True
                    
                    ' Verificar si el filtro se aplicó correctamente
                    If RsCabecera.EOF Then
                        For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                        MsgBox "No se encontró el cliente con el nombre: " & sDescrip, vbInformation, "Búsqueda"
                        Call cmdClearFilter_Click(cols)
                        RsCabecera.Filter = adFilterNone
                    Else
                        
                        For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                        ' Si se encuentran registros, accede al primer registro
                       RsCabecera.MoveFirst
                       
                       ' Cambia el valor en el Recordset (por ejemplo, en la columna "Cliente")
    '                   RsCabecera.Fields("Cliente").value = nuevoValor ' Sustituye "NuevoValor" con el valor que deseas asignar
                        RsCabecera.Update
    
                        grdGrilla.HoldFields    '
                        grdGrilla.col = c
                        grdGrilla.EditActive = True
                        
                        grdGrilla.Refresh
    '
                    End If
                Else
                    For Each col In cols
                            If col.ColIndex = 2 Then
                                col.FilterText = LCase(sDescrip)
                            End If
                        Next col
                    grdGrilla.HoldFields
                    RsCabecera.Filter = getFilter(cols)
                    grdGrilla.col = c
                    grdGrilla.EditActive = True
                End If
            End If
        Else
            grdGrilla.HoldFields
            RsCabecera.Filter = getFilter(cols)
            grdGrilla.col = c
            grdGrilla.EditActive = True
        End If
    Else
        grdGrilla.HoldFields
        RsCabecera.Filter = getFilter(cols)
        grdGrilla.col = c
        grdGrilla.EditActive = True
        
    End If

    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdGrilla
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0  'Detalle
                Sw = True
                'Cambiar el Nombre del Formulario Detalle
                frmReservaDetalle.Show vbModal
           
           Case Is = 1  'Detalle
                If RsCabecera.RecordCount > 0 Then
                
                    Dim ValorEstado As String
            
                    ValorEstado = IIf(IsNull(RsCabecera!EstadoReserva), "", RsCabecera!EstadoReserva)
                    
                    If UCase(ValorEstado) = "ATENDIDO" Or UCase(ValorEstado) = "ANULADO" Then
                        MsgBox "NO ES POSIBLE MODIFICAR UNA RESERVA QUE YA HA SIDO ATENDIDA O QUE HA SIDO ANULADA.", vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    Sw = False
                    'Cambiar el Nombre del Formulario Detalle
                    frmReservaDetalle.Show vbModal
                Else
                    MsgBox "No Existe Datos", vbExclamation, sMensaje
                End If
           Case Is = 2 'Crear Pedido de la reserva
                If RsCabecera.RecordCount > 0 Then
                    With frmReserva.RsCabecera
                      Dim cReserva As String
                      Dim eReserva As String
                      cReserva = IIf(IsNull(!TRESERVA), "", !TRESERVA)
                      eReserva = !tEstadoReserva
                      
                      If eReserva = "02" Or eReserva = "03" Then
                        MsgBox "SOLO SE PUEDE CREAR PEDIDO A LAS RESERVAS GENERADAS.", vbExclamation, sMensaje
                        Exit Sub
                      End If

                      Dim oComando As clsComando
                      Set oComando = New clsComando
                        If Not oComando.CreateCmdSp("SP_COVER_INSERTA_PEDIDO_ADICION", Cn) Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                    
                      oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                      oComando.CreateParameter "@codigoReserva", adVarChar, adParamInput, 10, cReserva
                      oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozo
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
                        MsgBox "Hubo un error en la generacion del pedido de la reserva: " & cReserva & vbNewLine & "Favor de volver a intentarlo o consultar a su administrador!!!", vbExclamation, sMensaje
                        Exit Sub
                      Else
                        frmCargoMozo.RsCabecera.Requery
                        frmCargoMozo.RsCabecera.MoveFirst
                        frmCargoMozo.RsCabecera.Find "Codigo = '" & sPedido & "'"
                        Unload Me
                      End If
                    End With
                Else
                    MsgBox "No Existe Datos.", vbExclamation, sMensaje
                End If
                
           Case Is = 3 'Salir
                Unload Me
           Case Is = 5  'Emite
                'frmRepReservas.Show vbModal
           Case Is = 6  'Salir
                'Unload Me
           
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ' Ingrese el Titulo
    Me.Caption = " Correlativo de Reservas "
    grdGrilla.Caption = Me.Caption
    
    Centrar Me
    nColumna = 0
    
    dtpFecIni.value = FechaServidor() - nDias
    dtpFecFin.value = FechaServidor()
    
    ' Ingrese la Vista
    Isql = "SELECT TRESERVA.tReserva,TRC.PedidoInfoRest, TRESERVA.fFecha, TRESERVA.fHora, TRESERVA.tApellido,TRESERVA.tCodigoDelivery, TDELIVERY.tApellido+' '+TDELIVERY.tNombre AS Cliente, TRESERVA.tNombre,TRESERVA.tMesero, TRESERVA.tMotivo, TRESERVA.tTelefono, TRESERVA.nPax, TRESERVA.tEstadoReserva, TRESERVA.tObservacion, TRESERVA.fRegistro, TRESERVA.tMesa, vEstadoReserva.Descripcion AS EstadoReserva, " & _
           "convert(varchar, fFecha, 103) as xFecha, convert(varchar, fHora, 103) as xHora, str(nPax,10,2) as xPAx " & _
           "FROM TRESERVA LEFT JOIN vEstadoReserva ON TRESERVA.tEstadoReserva = vEstadoReserva.Codigo LEFT JOIN TDELIVERY ON TRESERVA.tCodigoDelivery = TDELIVERY.tCodigoDelivery LEFT JOIN TRESERVA_COVERMANAGER TRC on TRESERVA.tReserva=TRC.CodigoReservaInfoRest " & _
           "where tPrioridad = '1' and tmesero in ('0000','" & sMozo & "' ) and fFecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' " & _
           "order by tReserva"
    
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    'Configuración de la Grilla
    Call ConfGrilla(9, grdGrilla, "Reserva", 2, "tReserva", 1200, 2, 0, "", _
                                  "Pedido", 2, "PedidoInforest", 1200, 2, 0, "", _
                                  "Nombres", 2, "Cliente", 3000, 0, 0, "", _
                                  "Teléfono", 2, "tTelefono", 1200, 0, 0, "", _
                                  "Fecha", 2, "xFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                  "Hora", 2, "xHora", 800, 0, 0, "", _
                                  "PAX", 2, "xPax", 800, 1, 0, "#,##0", _
                                  "Estado", 2, "EstadoReserva", 1200, 0, 0, "", _
                                  "Observación", 2, "tObservacion", 5500, 0, 0, "")
    
        
    Set grdGrilla.DataSource = RsCabecera
    RsCabecera.Filter = "tEstadoReserva ='01'"
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set rsReporte = Nothing
    Set RsCabecera = Nothing
    Set frmReserva = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub


Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
   Isql = "SELECT TRESERVA.tReserva,TRC.PedidoInfoRest, TRESERVA.fFecha, TRESERVA.fHora, TRESERVA.tApellido,TRESERVA.tCodigoDelivery, TDELIVERY.tApellido+' '+TDELIVERY.tNombre AS Cliente, TRESERVA.tNombre,TRESERVA.tMesero, TRESERVA.tMotivo, TRESERVA.tTelefono, TRESERVA.nPax, TRESERVA.tEstadoReserva, TRESERVA.tObservacion, TRESERVA.fRegistro, TRESERVA.tMesa, vEstadoReserva.Descripcion AS EstadoReserva, " & _
          "convert(varchar, fFecha, 103) as xFecha, convert(varchar, fHora, 103) as xHora, str(nPax,10,2) as xPAx " & _
          "FROM TRESERVA LEFT JOIN vEstadoReserva ON TRESERVA.tEstadoReserva = vEstadoReserva.Codigo LEFT JOIN TDELIVERY ON TRESERVA.tCodigoDelivery = TDELIVERY.tCodigoDelivery LEFT JOIN TRESERVA_COVERMANAGER TRC on TRESERVA.tReserva=TRC.CodigoReservaInfoRest " & _
          "where tPrioridad = '1' and tmesero in ('0000','" & sMozo & "' ) and fFecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' " & _
          "order by tReserva"
          
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub



