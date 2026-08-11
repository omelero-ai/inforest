VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCuentaCobrar 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8625
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   13725
   Icon            =   "frmCuentaCobrar.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595.96
   ScaleMode       =   0  'User
   ScaleWidth      =   13725
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   13665
      TabIndex        =   11
      Top             =   7245
      Width           =   13725
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
         Left            =   5580
         Picture         =   "frmCuentaCobrar.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Detalle del Documento"
         Top             =   90
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
         Index           =   0
         Left            =   8355
         Picture         =   "frmCuentaCobrar.frx":1884
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Anular Cuenta por Cobrar"
         Top             =   90
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
         Index           =   5
         Left            =   6135
         Picture         =   "frmCuentaCobrar.frx":1986
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Pagar el Documento"
         Top             =   90
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5490
         TabIndex        =   15
         Top             =   30
         Width           =   5550
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCuentaCobrar.frx":1A80
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
            Picture         =   "frmCuentaCobrar.frx":1FC2
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
            Picture         =   "frmCuentaCobrar.frx":2504
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4935
            Picture         =   "frmCuentaCobrar.frx":2A46
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4395
            Picture         =   "frmCuentaCobrar.frx":2F88
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3855
            Picture         =   "frmCuentaCobrar.frx":34CA
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
            Width           =   2190
         End
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
         Left            =   6690
         Picture         =   "frmCuentaCobrar.frx":3A0C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Emite"
         Top             =   90
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
         Left            =   8910
         Picture         =   "frmCuentaCobrar.frx":3F3E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   3840
         Picture         =   "frmCuentaCobrar.frx":4030
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filtrar"
         Top             =   700
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas"
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
         Left            =   135
         TabIndex        =   12
         Top             =   650
         Width           =   3525
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2070
            TabIndex        =   8
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
            Format          =   50724865
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
            Format          =   50724865
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
            TabIndex        =   14
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
            Left            =   1800
            TabIndex        =   13
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   555
         Index           =   9
         Left            =   7800
         Picture         =   "frmCuentaCobrar.frx":4132
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Reportes"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7245
         Picture         =   "frmCuentaCobrar.frx":4664
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar HTML"
         Top             =   90
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   45
      TabIndex        =   9
      Top             =   90
      Width           =   13560
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   405
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6870
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   13365
         _ExtentX        =   23574
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
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTipsWidth   =   0
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
   End
End
Attribute VB_Name = "frmCuentaCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim rsReporte As Recordset
Dim RsImpresion As Recordset
Dim Reporte As New dsrCuentaCobrar
Dim nColumna As Integer
Dim RsEmision As Recordset
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean
Dim xTabla As String
Dim nTotal As Double

Private Sub CmdEmite_Click()
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

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
   
   'Isql = "select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha from vDocumentoGrilla where fFecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
   'Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    Cn.Execute "delete from " & xTabla
    Cn.Execute "insert " & xTabla & " select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, Vencimiento, convert(bit,0) sel from vDocumentoGrilla " & _
    "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
    
   Isql = "select * from " & xTabla & " order by tdocumento"
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> "10" Then
    Cancel = True
End If
'MsgBox grdGrill.Columns(ColIndex).CellValue(11)

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
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & Err.Description, vbCritical + vbOKOnly, sMensaje

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



Public Function obtieneFilasSeleccionadas() As Integer
    obtieneFilasSeleccionadas = 0
    Dim m, i As Integer
    Dim k As Integer
    Dim primer As Integer
    primer = 0
    'Dim nTotal As Double
   ' nTotal = 0
    k = 0
'      For m = 1 To Me.grdGrilla.ApproxCount
'               If grdGrilla.Columns(10).CellText(m) = "-1" Then
'                    k = k + 1
'                    nTotal = nTotal + Val(grdGrilla.Columns(7).CellText(m))
'                End If
'        Next m
        
        RsCabecera.MoveFirst
        Do Until RsCabecera.EOF
            If (RsCabecera.Fields(16) = True) Then
                k = k + 1
                nTotal = nTotal + Val(RsCabecera.Fields(13))
                primer = RsCabecera.AbsolutePosition
                sPedido = RsCabecera!tCodigoPedido
            End If
        RsCabecera.MoveNext
        Loop
    RsCabecera.AbsolutePosition = IIf(primer = 0, 1, primer)
    obtieneFilasSeleccionadas = k
End Function


Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
          Case Is = 0  'Anular Pago
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoDocumento = "04" Then
                   MsgBox "Es un Documento Anulado", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoDocumento = "01" Then
                   MsgBox "Es un Documento Emitido", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If RsCabecera!tEstadoDocumento = "03" Then
                   'Password
                   If Supervisor("08") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                   'validactaporcobrar
                   
                   Dim codigoCliente As String
                   Dim nConsumoActual As Double
                   
                   codigoCliente = Calcular("select tclientepago as codigo from mdocumento where tdocumento='" & RsCabecera!tDocumento & "' ", Cn)
                   
                   nConsumoActual = Calcular("select isnull(nconsumoporcobrar,0) as codigo from tdelivery where tcodigodelivery='" & codigoCliente & "'", Cn)
                   
                   If nConsumoActual - Val(RsCabecera!nVenta) >= 0 Then
                   
                    Cn.Execute "update tdelivery set NCONSUMOPORCOBRAR=isnull(NCONSUMOPORCOBRAR,0) - " & Val(RsCabecera!nVenta) & " where tcodigodelivery='" & codigoCliente & "'"
                    Else
                    Cn.Execute "update tdelivery set NCONSUMOPORCOBRAR=0 where tcodigodelivery='" & codigoCliente & "'"
                    End If
                                      
                   'validactaporcobrar
                   
                   Cn.Execute "Update MDOCUMENTO set tClientePago = '', tEstadoDocumento = '01',lreplica=1 ,tcortesia='' where tDocumento = '" & RsCabecera!tDocumento & "'"
                   'Cn.Execute "Update MPEDIDO set tEstadoPedido = '01'  where tCodigoPedido in (select distinct tCodigoPedido from DPEDIDO where tDocumento='" & RsCabecera!tDocumento & "')"
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & RsCabecera!tDocumento & "'"

                   RsCabecera.Requery
                   If RsCabecera.RecordCount <> 0 Then
                      RsCabecera.MoveLast
                      cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
                   Else
                      cmdTexto.Caption = "Registro 0 de 0"
                   End If
                   Screen.MousePointer = vbDefault
                   Exit Sub
                End If
           
                If RsCabecera!tCaja <> sCaja Then
                   MsgBox "No se puede anular el pago porque pertenece a otra Caja (Caja " & sCaja & ") ", vbCritical, sMensaje
                   Exit Sub
                End If
                'grdGrilla
                If RsCabecera!tTurno <> sTurno Then
                   MsgBox "No se puede Anular el Pago de un documento de otro turno", vbCritical, sMensaje
                   Exit Sub
                End If
                
                If MsgBox("Seguro de Anular el Pago? " & Format(grdGrilla.Columns(2).Text, "@-@@@@@-@@@@@@@@@") & _
                          "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                Cn.Execute "delete from DPAGODOCUMENTO where tDocumento = '" & RsCabecera!tDocumento & "'"
                Cn.Execute "UPDATE VC " _
                            & "SET VC.lESTADO_USO = 0 " _
                            & "FROM DPAGODOCUMENTO_VC PVC " _
                            & "INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                            & "WHERE PVC.tDocumento = '" & RsCabecera!tDocumento & "' AND " _
                            & "CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "

                Cn.Execute "DELETE PVC " _
                            & " FROM DPAGODOCUMENTO_VC PVC " _
                            & " INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                            & " WHERE PVC.tDocumento = '" & RsCabecera!tDocumento & "' AND " _
                            & " CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
                                
                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '01',lreplica=1  where tDocumento = '" & RsCabecera!tDocumento & "'"
                Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & RsCabecera!tDocumento & "'"
                'Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', tMesa =''  where tCodigoPedido ='" & RsCabecera!tCodigoPedido & "'"
                   
                RsCabecera.Requery
                If RsCabecera.RecordCount <> 0 Then
                   RsCabecera.MoveLast
                   cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
                Else
                   cmdTexto.Caption = "Registro 0 de 0"
                End If
                Screen.MousePointer = vbDefault
           
           Case Is = 1  'Detalle
                If RsCabecera.RecordCount > 0 Then
                   sDocumento = grdGrilla.Columns(2).Text
                   frmCuentaCobrarDetalle.Show vbModal
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
           
           Case Is = 5  'Pagar
                
                If RsCabecera.RecordCount > 0 Then
                    Dim t As Integer
                    t = obtieneFilasSeleccionadas
                    If t = 0 Then
                        MsgBox "Debe seleccionar Documentos ", vbInformation, sMensaje
                        Exit Sub
                    Else
                        If t = 1 Then
                        
                            If RsCabecera!tEstadoDocumento = "03" Then
                               nCargo = Round(RsCabecera!nVenta, 2)
                               sDocumento = RsCabecera!tDocumento
                               xTipo = "PorPagar"
                               sFormulario = ""
                               frmPago.cmdOpcion(8).Enabled = False
                               frmPago.Show vbModal
                               
                                Cn.Execute "delete from " & xTabla
                                Cn.Execute "insert " & xTabla & " select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, Vencimiento, convert(bit,0) sel from vDocumentoGrilla " & _
                                "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
                                
                                Isql = "select * from " & xTabla & " order by tdocumento"
                                Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
                                Set grdGrilla.DataSource = RsCabecera
                                cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
                                nTotal = 0
                               'RsCabecera.Requery
                             Else
                               MsgBox "No Es un Documento por Cancelar", vbExclamation, sMensaje
                             End If
                        
                        
                        Else
                            frmPagoMasivo.obtieneSaldo nTotal
                            frmPagoMasivo.Show vbModal
                            If wEnter = True Then
                                        Dim b As Integer
                                        Dim pDocumento As String
                                        'Cuentas x Cobrar
                                        Dim clientePagoCP As String
                                        Dim totalDocumentoCP As Double
                                        Dim nConsumoActualCP As Double
                                        Dim totalCP As Double
                                        Dim nTotalDoc As Double
                                        RsCabecera.MoveFirst
                                        Do Until RsCabecera.EOF
                                        'For b = 1 To Me.grdGrilla.ApproxCount
                                            'If grdGrilla.Columns(10).CellValue(b) = True Then
                                            If (RsCabecera.Fields(16) = True) Then
                                                clientePagoCP = ""
                                                totalDocumentoCP = 0
                                                nConsumoActualCP = 0
                                                totalCP = 0
                                                pDocumento = RsCabecera.Fields(3) 'Replace(grdGrilla.Columns(2).CellValue(b), "-", "")
                                                nTotalDoc = Val(RsCabecera.Fields(13)) 'Val(grdGrilla.Columns(7).CellValue(b))
                                               
                                                If tTipoPagoMasivo = "E" Then
                                                    'efectivo
                                                    Isql = "insert into DPAGODOCUMENTO " & _
                                                            "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, fRegistro, tUsuario,fDiaContable, tForma ) " & _
                                                            "Values(   '" & pDocumento & "','1','" & sTurno & "','01','01', " & nTC & ", " & nTotalDoc & ",  getdate(),'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','PM')"
                                                    Cn.Execute Isql
                                                End If
                                                
                                                If tTipoPagoMasivo = "C" Then
                                                    'Cheque
                                                    Isql = "insert into DPAGODOCUMENTO " & _
                                                           "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, tBanco, tNumero, fRegistro, tUsuario,fDiaContable,tForma ) " & _
                                                           "Values (   '" & pDocumento & "','1','" & sTurno & "','03','01'," & nTC & "," & nTotalDoc & ",'" & tBancoPagoMasivo & "','" & tDocumentoPagoMasivo & "', getdate(),'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','PM')"
                                                    
                                                    Cn.Execute Isql
                                                End If
                    
                                                If tTipoPagoMasivo = "O" Then
                                                    'documentos
                                                    Isql = "insert into DPAGODOCUMENTO " & _
                                                           "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable,tForma ) " & _
                                                           "Values(   '" & pDocumento & "', '1', '" & sTurno & "','04', '" & tOtroTipoCancelacionMasivo & "', '01'," & nTC & ", " & nTotalDoc & ", '" & tDocumentoPagoMasivo & "',  '" & tTipoTarjetaMasivo & "', getdate(),'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','PM')"
                                                    Cn.Execute Isql
                                                    
                                                    If tOtroTipoCancelacionMasivo = "001" Then
                                                       Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & tDocumentoPagoMasivo & "'"
                                                    ElseIf tOtroTipoCancelacionMasivo = "002" Then
                                                       Cn.Execute "update MNOTACREDITO set tEstadoDocumento ='02',lreplica=1 where tNotaCredito ='" & tDocumentoPagoMasivo & "'"
                                                    End If
                                                End If
                                                                      
                                                If tTipoPagoMasivo = "T" Then
                                                    'tarjetas
                                                    Isql = "insert into DPAGODOCUMENTO " & _
                                                           "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable, tForma ) " & _
                                                           "Values(    '" & pDocumento & "','1','" & sTurno & "','02','01','', " & nTC & ", " & nTotalDoc & ", 0,'" & tTipoTarjetaMasivo & "', '" & tDocumentoPagoMasivo & "', '" & tOtroTipoCancelacionMasivo & "',getdate(),'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','PM')"
                                                 
                                                    Cn.Execute Isql
                                                End If
                                            
                                                If tTipoPagoMasivo <> "" Then
                                                    Isql = "Update MDOCUMENTO set " & _
                                                            "tEstadoDocumento = '02', " & _
                                                            "nAbono = " & nTotalDoc & ", " & _
                                                             "fPago = getdate()," & _
                                                            "nVuelto = 0 " & _
                                                            ",lreplica=1, TCORTESIA=''  where tDocumento ='" & pDocumento & "'"
                                    
                                                   Cn.Execute Isql
                                           
                                                    clientePagoCP = Calcular("Select tClientePago As Codigo From MDOCUMENTO Where tDocumento = '" & pDocumento & "'", Cn)
                                                    nConsumoActualCP = Calcular("Select isnull(nconsumoporcobrar,0) as codigo From TDELIVERY where tCodigoDelivery='" & clientePagoCP & "'", Cn)
                                                    totalDocumentoCP = Calcular("Select nVenta As Codigo From mdocumento where tdocumento = '" & pDocumento & "'", Cn)
                                                    totalCP = nConsumoActualCP - totalDocumentoCP
                                                    Isql = "Update TDELIVERY Set nConsumoPorCobrar = " & totalCP & " Where tCodigoDelivery='" & clientePagoCP & "'"
                                                    Cn.Execute Isql
                                                    
                                                End If

                                            End If
                                            RsCabecera.MoveNext
                                        Loop
                                        'Next b
                            
                                        tTipoPagoMasivo = ""
                                        nTotal = 0
'                                        Cn.Execute "delete from " & xTabla
'                                        Cn.Execute "insert " & xTabla & " select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, Vencimiento, convert(bit,0) sel from vDocumentoGrilla " & _
'                                        "where fFecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
'
'                                        Isql = "select * from " & xTabla & " order by tdocumento"
'                                        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
'                                        Set grdGrilla.DataSource = RsCabecera
'                                        cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
                                        
                            Else
                                nTotal = 0
                            End If
                            
                            Cn.Execute "delete from " & xTabla
                            Cn.Execute "insert " & xTabla & " select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, Vencimiento, convert(bit,0) sel from vDocumentoGrilla " & _
                            "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
                            
                            Isql = "select * from " & xTabla & " order by tdocumento"
                            Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
                            Set grdGrilla.DataSource = RsCabecera
                            cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
                                        
                        End If
                        
                    
                    End If
                    
                End If
                
           Case Is = 6  'Salir
                Unload Me
           
           Case Is = 8  'Cambio Documento
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoDocumento <> "02" Then
                   MsgBox "No Es un Documento Cancelado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If MsgBox("Seguro de Cambiar el Documento? " & Format(grdGrilla.Columns(2).Text, "@-@@@@@-@@@@@@@@@") & _
                          "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                sCodigo = RsCabecera!tCodigoCliente
                sDescrip = RsCabecera!Cliente
                sDocumento = RsCabecera!tDocumento
                frmCambio.Show vbModal
                RsCabecera.Requery
                
           Case Is = 9  'Reimpresion
                If RsCabecera.RecordCount = 0 Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Set rsReporte = RsCabecera.Clone
                rsReporte.Sort = grdGrilla.Columns(nColumna).DataField & " ASC"
                
                Reporte.DiscardSavedData
                Reporte.Database.SetDataSource rsReporte
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.Text15.SetText sRazonSocial
                Reporte.Text3.SetText sRazonSocial
                
                frmEmite.CRViewer.ReportSource = Reporte
                frmEmite.CRViewer.DisplayGroupTree = False
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
          
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    cmdOpcion(0).Enabled = IIf(wInicio, True, False)
    cmdOpcion(5).Enabled = IIf(wInicio, True, False)
    'Ingrese el Titulo
    Me.Caption = " Documentos Pagados "
    
    Centrar Me
    nColumna = 0
       dtpFecIni.Value = FechaServidor() - nDias
    dtpFecFin.Value = FechaServidor()
    
'        xTabla = dbTemporal(sCaja, 16, "tCaja", "nVarChar(3)", _
'                                    "tTurno", "nVarChar(15)", _
'                                    "tCodigoPedido", "nVarChar(10)", _
'                                    "tDocumento", "nVarChar(20)", _
'                                    "fFecha", "datetime", _
'                                    "tUsuario", "nVarChar(50)", _
'                                    "tCodigoCliente", "nVarChar(10)", _
'                                    "Cliente", "nVarChar(255)", _
'                                    "nventa", "Float", _
'                                    "EstadoDocumento", "nVarChar(50)", _
'                                    "tEstadoDocumento", "nVarChar(5)", _
'                                    "Cortesia", "nVarChar(100)", _
'                                    "ClientePago", "nVarChar(255)", _
'                                    "xVenta", "nVarChar(20)", _
'                                    "xFecha", "nVarChar(30)", _
'                                    "Sel", "bit")

        xTabla = dbTemporal(sCaja, 17, "tCaja", "nVarChar(3)", _
                                    "tTurno", "nVarChar(15)", _
                                    "tCodigoPedido", "nVarChar(10)", _
                                    "tDocumento", "nVarChar(20)", _
                                    "fFecha", "datetime", _
                                    "tUsuario", "nVarChar(50)", _
                                    "tCodigoCliente", "nVarChar(10)", _
                                    "Cliente", "nVarChar(255)", _
                                    "nventa", "Float", _
                                    "EstadoDocumento", "nVarChar(50)", _
                                    "tEstadoDocumento", "nVarChar(5)", _
                                    "Cortesia", "nVarChar(100)", _
                                    "ClientePago", "nVarChar(255)", _
                                    "xVenta", "nVarChar(20)", _
                                    "xFecha", "nVarChar(30)", _
                                    "Vencimiento", "nVarChar(50)", _
                                    "Sel", "bit")
                                                                                                                                 
    
    'Ingrese la Vista
    Cn.Execute "insert " & xTabla & " select tCaja, tTurno, tCodigoPedido, tDocumento, fFecha, tUsuario, tCodigoCliente, Cliente, nVenta, EstadoDocumento, tEstadoDocumento, Cortesia, ClientePago, str(nVenta,10,2) as xVenta, convert(varchar, fFecha, 103) as xFecha, Vencimiento, convert(bit,0) sel from vDocumentoGrilla " & _
    "where fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' and tEstadoDocumento = '03' order by tDocumento"
    
     Isql = "select * from " & xTabla & " order by tdocumento"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    nColumna = 0
    
    'Configuración de la Grillahhhh
'    Call ConfGrilla(10, grdGrilla, "Caja", 2, "tCaja", 500, 0, 0, "", _
'                                  "Turno", 2, "tTurno", 1200, 0, 0, "", _
'                                  "Documento", 2, "tDocumento", 2300, 0, 0, "@-@@@@@-@@@@@@@@@", _
'                                  "Cliente", 2, "Cliente", 3000, 0, 0, "", _
'                                  "Fec.Emis", 2, "xFecha", 1200, 0, 0, "dd/mm/yyyy", _
'                                  "Usu.Emis", 2, "tUsuario", 1200, 0, 0, "", _
'                                  "A Cuenta de", 2, "ClientePago", 3300, 0, 0, "", _
'                                  "Monto", 2, "xVenta", 1200, 1, 0, "###,###,##0.00", _
'                                  "Estado", 2, "EstadoDocumento", 1300, 0, 0, "", _
'                                  "Sel", 2, "Sel", 850, 2, 4, "")
    
    Call ConfGrilla(11, grdGrilla, "Caja", 2, "tCaja", 500, 0, 0, "", _
                                  "Turno", 2, "tTurno", 1200, 0, 0, "", _
                                  "Documento", 2, "tDocumento", 2300, 0, 0, "@-@@@@@-@@@@@@@@@", _
                                  "Cliente", 2, "Cliente", 3000, 0, 0, "", _
                                  "Fec.Emis", 2, "xFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                  "Usu.Emis", 2, "tUsuario", 1200, 0, 0, "", _
                                  "A Cuenta de", 2, "ClientePago", 3300, 0, 0, "", _
                                  "Monto", 2, "xVenta", 1200, 1, 0, "###,###,##0.00", _
                                  "Estado", 2, "EstadoDocumento", 1300, 0, 0, "", _
                                  "Vencimiento", 2, "Vencimiento", 1500, 2, 0, "", _
                                  "Sel", 2, "Sel", 850, 2, 4, "")
                                    
    Set grdGrilla.DataSource = RsCabecera
    
    'Resalta los cierres que estan en estado finalizado
    Call FormatoCeldaGrilla(grdGrilla, 9, "VENCIDO")
    
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsReporte = Nothing
    Set RsImpresion = Nothing
    Set RsEmision = Nothing
    Set RsCabecera = Nothing
    Set frmDocumentoCorrelativo = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
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
End Function
'diaContable

