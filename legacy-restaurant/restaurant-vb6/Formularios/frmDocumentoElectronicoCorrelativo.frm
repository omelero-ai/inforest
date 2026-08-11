VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDocumentoElectronicoCorrelativo 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8535
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   9990
   Icon            =   "frmDocumentoElectronicoCorrelativo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8506.263
   ScaleMode       =   0  'User
   ScaleWidth      =   9990
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9975
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6870
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   9735
         _ExtentX        =   17171
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
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
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         Left            =   135
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   9930
      TabIndex        =   2
      Top             =   7155
      Width           =   9990
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Resumen Diario"
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
         Index           =   3
         Left            =   7320
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Confirma Entrega"
         Top             =   45
         Visible         =   0   'False
         Width           =   1605
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
         Index           =   8
         Left            =   6170
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":1884
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Reenviar Documento Electronico"
         Top             =   40
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
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   5235
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   3195
            TabIndex        =   16
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   54919169
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   720
            TabIndex        =   17
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   54919169
            CurrentDate     =   37539
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
            Left            =   2700
            TabIndex        =   19
            Top             =   300
            Width           =   225
            WordWrap        =   -1  'True
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
            Left            =   90
            TabIndex        =   18
            Top             =   300
            Width           =   405
            WordWrap        =   -1  'True
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
         Index           =   6
         Left            =   6720
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":1CC6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5325
         TabIndex        =   5
         Top             =   30
         Width           =   5385
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":1DB8
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":22FA
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":283C
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4770
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":2D7E
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":32C0
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3690
            Picture         =   "frmDocumentoElectronicoCorrelativo.frx":3802
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   12
            Top             =   150
            Width           =   2025
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
         Left            =   5610
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":3D44
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Emite"
         Top             =   45
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
         Index           =   14
         Left            =   17685
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":4276
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   -450
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   5430
         Picture         =   "frmDocumentoElectronicoCorrelativo.frx":4368
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Filtrar"
         Top             =   710
         Width           =   555
      End
   End
   Begin VB.Image imageHash 
      Height          =   495
      Left            =   10920
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   375
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmDocumentoElectronicoCorrelativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsImpresion As Recordset
Dim rstFuente As ADODB.Recordset
Dim nColumna As Integer
Dim RsEmision As Recordset
Dim sUsuarioAutoriza As String
Dim flagTipoCambio As Boolean
Dim tTipoPago As String
Dim tnumerorecibo As String
Dim nRespuesta As Integer
Dim lImprimeAlternativa As Boolean
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Dim lPagoPV As Boolean

'FACTURACION_E_PERU
Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim lcodigoHash As Boolean
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String

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

Private Sub CmdOpcion_Click(Index As Integer)
    On Error GoTo fin
    Select Case Index
            Case Is = 3
            Dim MensajeTemp As String
            MensajeTemp = "El resumen Diario de Bajas se Generara en Base a la Fecha Final Ingresada en la parte Inferior del Formulario!!" & vbNewLine & "Se generar el Resumen Diario de la Fecha: " & Format(Me.dtpFecFin.value, "dd/mm/YYYY") & vbNewLine & "¿Esta seguro de Continuar?"
            If MsgBox(MensajeTemp, vbYesNo, sMensaje) = vbYes Then
                Screen.MousePointer = vbHourglass
                If Not INSERTA_FE_INFOREST("", 3, Me.dtpFecFin.value) Then
                    'MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                    Screen.MousePointer = vbDefault
                    'Exit Sub
                End If
                If Not INSERTA_FE_INFOREST("", 4, Me.dtpFecFin.value) Then
                    'MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                    Screen.MousePointer = vbDefault
                    'Exit Sub
                End If
                MsgBox "Proceso Concluido!!!", vbInformation, sMensaje
                RsCabecera.Requery
            End If
            
            Case Is = 6  'Salir
                Unload Me
            Case Is = 8 ' reenvio a paperlees
           
                If lFEpape Then
                    If RsCabecera.RecordCount = 0 Then
                       Exit Sub
                    End If
                    
                    If MsgBox("Seguro de Reenviar el Documento al E-Pos? " & Format(grdGrilla.Columns(1).value, "@-@@@@@-@@@@@@@@@") & _
                              "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If
                    
    '                If Supervisor("05") = False Then
    '                   MsgBox "Clave no permitida", vbExclamation, sMensaje
    '                   Exit Sub
    '                End If
                    
                    If (Calcular("select isnull(lFacturacionE,0) as codigo from vTipoDocumento where Codigo='" & RsCabecera!tipodocu & "'", Cn)) = False Then
                         MsgBox "ESTE DOCUMENTO NO ESTA REFERENCIADO COMO ELECTRONICO EN ESTA CAJA, FAVOR DE REVISAR LAS CONFIGURACIONES EN ADMINISTRACION!!", vbInformation, sMensaje
                         Exit Sub
                    End If
                    
                    Screen.MousePointer = vbHourglass
                    
                    If RsCabecera!razonsocial = "NOTACREDITO" Then
                        If pais = "000" And Not FacturarTCPIP(2, RsCabecera!nro_efact, 1) Then
                            MsgBox "Error al Enviar la Trama a E-Pos", vbInformation, sMensaje
                            Exit Sub
                        End If
                        
                        '-----------------------
                        If pais = "000" And lFEpape Then
                             If Not FacturarTCPIP(3, RsCabecera!nro_efact, 1) Then
                                MsgBox ("La confirmación ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                             End If
                        End If
                    Else
                        If pais = "000" And Not FacturarTCPIP(2, RsCabecera!nro_efact, 0) Then
                            MsgBox "Error al Enviar la Trama a E-Pos", vbInformation, sMensaje
                            Exit Sub
                        End If
                        
                        '-----------------------
                        If pais = "000" And lFEpape Then
                             If Not FacturarTCPIP(3, RsCabecera!nro_efact, 0) Then
                                MsgBox ("La confirmación ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                             End If
                        End If
                    End If
                    
                    RsCabecera.Requery
                ElseIf lFEBiz Then
                
                    If RsCabecera.RecordCount = 0 Then
                       Exit Sub
                    End If
                    
                    If MsgBox("Seguro de Enviar el Documento Electronico? " & Format(grdGrilla.Columns(1).value, "@-@@@@@-@@@@@@@@@") & _
                            "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If

                    If (Calcular("select isnull(lFacturacionE,0) as codigo from vTipoDocumento where Codigo='" & RsCabecera!tipodocu & "'", Cn)) = False Then
                         MsgBox "ESTE DOCUMENTO NO ESTA REFERENCIADO COMO ELECTRONICO EN ESTA CAJA, FAVOR DE REVISAR LAS CONFIGURACIONES EN ADMINISTRACION!!", vbInformation, sMensaje
                         Exit Sub
                    End If
                    
                   Screen.MousePointer = vbHourglass
                    If RsCabecera!Doc = "D" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mdocumento where tdocumento='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Exit Sub
                        End If
                        
                        If Not INSERTA_FE_INFOREST(RsCabecera!nro_efact, 1, DateTime.Date) Then
                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                    ElseIf RsCabecera!Doc = "N" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mNotacredito where tnotacredito='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Exit Sub
                        End If
                        If Not INSERTA_FE_INFOREST(RsCabecera!nro_efact, 2, DateTime.Date) Then
                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    MsgBox "Se realizo correctamente el envió!!!", vbInformation, sMensaje
                ElseIf lFEGesa Then
                
                    If RsCabecera.RecordCount = 0 Then
                       Exit Sub
                    End If
                    
                    If MsgBox("Seguro de Enviar el Documento Electronico? " & Format(grdGrilla.Columns(1).value, "@-@@@@@-@@@@@@@@@") & _
                            "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If

                    If (Calcular("select isnull(lFacturacionE,0) as codigo from vTipoDocumento where Codigo='" & RsCabecera!tipodocu & "'", Cn)) = False Then
                         MsgBox "ESTE DOCUMENTO NO ESTA REFERENCIADO COMO ELECTRONICO EN ESTA CAJA, FAVOR DE REVISAR LAS CONFIGURACIONES EN ADMINISTRACION!!", vbInformation, sMensaje
                         Exit Sub
                    End If
                    
                   Screen.MousePointer = vbHourglass
                    If RsCabecera!Doc = "D" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mdocumento where tdocumento='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        If Not INSERTAFE(RsCabecera!nro_efact, "", 1, "") Then  '----CABECERA
                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If

                    ElseIf RsCabecera!Doc = "N" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mNotacredito where tnotacredito='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        Dim oComando As clsComando
                        Set oComando = New clsComando
                        If Not oComando.CreateCmdSp("USP_FactNotaCredito", Cn) Then
                             Set oComando = Nothing
                             Screen.MousePointer = vbDefault
                             Exit Sub
                        End If
                        oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, RsCabecera!nro_efact
                        oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 250, ""
                        If Not oComando.GetParamOK Then
                             Set oComando = Nothing
                             Screen.MousePointer = vbDefault
                             Exit Sub
                        End If
                        If Not oComando.ExecSP Then
                             Set oComando = Nothing
                             Screen.MousePointer = vbDefault
                             Exit Sub
                        End If
'                        If Not INSERTA_FE_INFOREST(RsCabecera!nro_efact, 2, DateTime.Date) Then
'                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
'                            Screen.MousePointer = vbDefault
'                            Exit Sub
'                        End If
                    End If
                    MsgBox "Se realizo correctamente el Envió!!!", vbInformation, sMensaje
                Else
                 If RsCabecera.RecordCount = 0 Then
                       Exit Sub
                    End If
                    
                    If MsgBox("Seguro de Enviar el Documento Electronico? " & Format(grdGrilla.Columns(1).value, "@-@@@@@-@@@@@@@@@") & _
                            "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If

                    If (Calcular("select isnull(lFacturacionE,0) as codigo from vTipoDocumento where Codigo='" & RsCabecera!tipodocu & "'", Cn)) = False Then
                         MsgBox "ESTE DOCUMENTO NO ESTA REFERENCIADO COMO ELECTRONICO EN ESTA CAJA, FAVOR DE REVISAR LAS CONFIGURACIONES EN ADMINISTRACION!!", vbInformation, sMensaje
                         Exit Sub
                    End If
                    
                   Screen.MousePointer = vbHourglass
                    If RsCabecera!Doc = "D" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mdocumento where tdocumento='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        If Not INSERTAFE(RsCabecera!nro_efact, "", 1, "") Then  '----CABECERA
                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If

                    ElseIf RsCabecera!Doc = "N" Then
                        If Calcular("select isnull(lestadofacturacion,0) as codigo from mNotacredito where tnotacredito='" & RsCabecera!nro_efact & "'", Cn) = True Then
                            MsgBox "El documento ya fue enviado, Verificar en el portal!!", vbInformation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        If Not INSERTAFE(RsCabecera!nro_efact, "", 2, "") Then  '----CABECERA
                            MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    MsgBox "Se realizo correctamente el envió!!!", vbInformation, sMensaje
                    'MsgBox "No hay activaciones electronicas para reenvio, favor de verificar las configuraciones!!!!"
                End If
                
           RsCabecera.Requery
            'Screen.MousePointer = vbDefault
    End Select
     
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & error
    RsCabecera.Requery
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
    If lFEpape Then
        Isql = "select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mdocumento where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,'" & Format(dtpFecIni.value, "yyyy/MM/dd") & "') and convert(date,'" & Format(dtpFecFin.value, "yyyy/MM/dd") & "') " & _
               " union select tcaja as caja,tnotacredito as nro_efact ,fregistro,ttipodocumento as tipodocu, 'NOTACREDITO' as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mnotacredito where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,'" & Format(dtpFecIni.value, "yyyy/MM/dd") & "') and convert(date,'" & Format(dtpFecFin.value, "yyyy/MM/dd") & "') "
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    ElseIf lFEBiz Or lFEGesa Then
        Isql = "exec usp_ListDocumentosFE '" & sCaja & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & "',2 "
        'Isql = "select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mdocumento where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,'" & Format(dtpFecIni.value, "yyyy/MM/dd") & "') and convert(date,'" & Format(dtpFecFin.value, "yyyy/MM/dd") & "') and tcaja='" & sCaja & "' "
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    Else
        Isql = "exec usp_ListDocumentosFE '" & sCaja & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & "',2 "
        'Isql = "select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mdocumento where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,'" & Format(dtpFecIni.value, "yyyy/MM/dd") & "') and convert(date,'" & Format(dtpFecFin.value, "yyyy/MM/dd") & "') and tcaja='" & sCaja & "' "
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    End If

   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub


Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    dtpFecIni.value = FechaServidor() - nDias
    dtpFecFin.value = FechaServidor()
      
    'Ingrese el Titulo
    Me.Caption = " Documentos Electrónicos "
    
    Centrar Me
    nColumna = 0
   
   
    'Ingrese la Vista
    If lFEpape Then
        Isql = "select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mdocumento where ttipodocumento<>'00' and convert(date,fregistro) = convert(date,getdate()) and tcaja='" & sCaja & "'" & _
               " UNION select tcaja as caja,tnotaCredito as nro_efact ,fregistro,ttipodocumento as tipodocu, 'NOTACREDITO' as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mnotacredito where ttipodocumento<>'00' and convert(date,fregistro) = convert(date,getdate()) and tcaja='" & sCaja & "'"
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    ElseIf lFEBiz Or lFEGesa Or lFETusFacturasAPP Then
        Isql = "exec usp_ListDocumentosFE '" & sCaja & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & "',2 "
        'Isql = "select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial,lestadofacturacion as cdr, Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, tusuario as cajero from dbo.mdocumento where ttipodocumento<>'00' and convert(date,fregistro) = convert(date,getdate()) and tcaja='" & sCaja & "'"
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    Else
       Isql = "exec usp_ListDocumentosFE '" & sCaja & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & "',2 "
        Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
'        Isql = "select caja,nro_efact,fregistro,tipodocu,razonsocial,cdr,Case When ISNULL(cdr,'')  = '' then 'Documento en Proceso' else descripcionCDR end As cdrDes, numerorefe, cajero from dbo.DOCUMENTOVENTA"
'        Set RsCabecera = Lib.OpenRecordset(Isql, CnFE)
    End If

     nColumna = 0
    
      If lFEpape Then
           'Configuración de la Grilla
            Call ConfGrilla(7, grdGrilla, "Caja", 2, "caja", 600, 2, 0, "", _
                                    "Documento", 2, "nro_efact", 1600, 0, 0, "", _
                                    "Fec.Emis", 2, "fregistro", 1000, 0, 0, "dd/mm/yyyy", _
                                    "Tipo", 2, "tipodocu", 800, 2, 0, "", _
                                    "Cliente", 2, "razonsocial", 3200, 0, 0, "", _
                                    "Estado ", 2, "cdrDes", 2500, 0, 0, "", _
                                    "Doc.Referencia", 2, "numerorefe", 1400, 0, 0, "", _
                                    "Usu.Emis", 2, "cajero", 1500, 0, 0, "")
    ElseIf lFEBiz Or lFEGesa Then
           'Configuración de la Grilla
            Call ConfGrilla(10, grdGrilla, "Caja", 2, "caja", 600, 2, 0, "", _
                                    "Documento", 2, "nro_efact", 1800, 0, 0, "", _
                                    "Fec.Emis", 2, "fregistro", 1000, 0, 0, "dd/mm/yyyy", _
                                    "Tipo", 2, "tipodocu", 800, 2, 0, "", _
                                    "Cliente", 2, "razonsocial", 3200, 0, 0, "", _
                                    "Estado ", 2, "cdrDes", 1500, 0, 0, "", _
                                    "Doc.Referencia", 2, "numerorefe", 1600, 0, 0, "", _
                                    "Descripcion", 2, "DocDescripcion", 2000, 0, 0, "", _
                                    "Usu.Emis", 2, "cajero", 1500, 0, 0, "", _
                                    "Respuesta", 2, "Firma", 5000, 0, 0, "")

      Else
           'Configuración de la Grilla
'            Call ConfGrilla(8, grdGrilla, "Caja", 2, "caja", 600, 2, 0, "", _
'                                    "Documento", 2, "nro_efact", 1400, 0, 0, "", _
'                                    "Fec.Emis", 2, "fregistro", 1000, 0, 0, "dd/mm/yyyy", _
'                                    "Tipo", 2, "tipodocu", 800, 2, 0, "", _
'                                    "Cliente", 2, "razonsocial", 3200, 0, 0, "", _
'                                    "CDR", 2, "cdr", 650, 2, 0, "", _
'                                    "Estado CDR", 2, "cdrDes", 4000, 0, 0, "", _
'                                    "Doc.Referencia", 2, "numerorefe", 1400, 0, 0, "", _
'                                    "Usu.Emis", 2, "cajero", 1500, 0, 0, "")

            Call ConfGrilla(10, grdGrilla, "Caja", 2, "caja", 600, 2, 0, "", _
                                    "Documento", 2, "nro_efact", 1800, 0, 0, "", _
                                    "Fec.Emis", 2, "fregistro", 1000, 0, 0, "dd/mm/yyyy", _
                                    "Tipo", 2, "tipodocu", 800, 2, 0, "", _
                                    "Cliente", 2, "razonsocial", 3200, 0, 0, "", _
                                    "Estado ", 2, "cdrDes", 1500, 0, 0, "", _
                                    "Doc.Referencia", 2, "numerorefe", 1600, 0, 0, "", _
                                    "Descripcion", 2, "DocDescripcion", 2000, 0, 0, "", _
                                    "Usu.Emis", 2, "cajero", 1500, 0, 0, "", _
                                    "Respuesta", 2, "Firma", 7000, 0, 0, "")

      End If
      
    If lFeGoodHope Then CmdOpcion(3).Visible = False
    If lFEpape Then CmdOpcion(3).Visible = False
    If lFEGesa Then CmdOpcion(3).Visible = False
    
    If nTC = 0 Then: flagTipoCambio = True
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsImpresion = Nothing
    Set RsEmision = Nothing
    Set RsCabecera = Nothing
    Set frmDocumentoCorrelativo = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub


