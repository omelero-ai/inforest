VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCaja 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8235
   ClientLeft      =   2535
   ClientTop       =   1725
   ClientWidth     =   13740
   Icon            =   "frmCaja.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   13740
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   7440
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   13695
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6915
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12197
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   13680
      TabIndex        =   0
      Top             =   7530
      Width           =   13740
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7650
         Picture         =   "frmCaja.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exportar HTML"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   555
         Index           =   5
         Left            =   8235
         Picture         =   "frmCaja.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Reportes"
         Top             =   90
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
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
         Height          =   645
         Left            =   9405
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   3660
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2205
            TabIndex        =   15
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
            Format          =   56492033
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   16
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
            Format          =   56492033
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
            Left            =   1845
            TabIndex        =   18
            Top             =   270
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
            Left            =   0
            TabIndex        =   17
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdProcesa 
         Enabled         =   0   'False
         Height          =   555
         Left            =   13095
         Picture         =   "frmCaja.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Filtrar"
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
         Picture         =   "frmCaja.frx":0FA8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
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
         Picture         =   "frmCaja.frx":109A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Index           =   1
         Left            =   6480
         Picture         =   "frmCaja.frx":15CC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Modifica"
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
         Left            =   5895
         Picture         =   "frmCaja.frx":16CE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Agrega"
         Top             =   90
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   1
         Top             =   30
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmCaja.frx":1C00
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmCaja.frx":2142
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmCaja.frx":2684
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCaja.frx":2BC6
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmCaja.frx":3108
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCaja.frx":364A
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   8
            Top             =   150
            Width           =   2505
         End
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsReporte As Recordset
Dim nColumna As Integer
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Private Sub cmdEmite_Click()
    With grdGrilla.PrintInfo
        'Cabecera
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
   fraGrilla.Height = IIf(Me.Height - 1300 > 0, Me.Height - 1300, 0)
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

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0  'Nuevo
                Sw = True
                'Cambiar el Nombre del Formulario Detalle
                frmCajaDetalle.Show vbModal
                      
           Case Is = 1  'Modifica
                If RsCabecera.RecordCount > 0 Then
                    Sw = False
                    'Cambiar el Nombre del Formulario Detalle
                    frmCajaDetalle.Show vbModal
                Else
                    MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
                               
           Case Is = 6  'Salir
                Unload Me
           
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
 
    ' Ingrese el Titulo
    Me.Caption = " Mantenimiento de Cajas "
    grdGrilla.Caption = Me.Caption
    Centrar Me
    nColumna = 0
    
    'Ingrese la Vista
    Isql = " SELECT dbo.TCAJA.tCaja, dbo.TCAJA.tDescripcion, dbo.TCAJA.tPrecuenta, dbo.TCAJA.lComanda, dbo.TCAJA.vComanda, dbo.TCAJA.lMotivoEliminaC, dbo.TCAJA.lMotivoElimina, dbo.TCAJA.lActivo, dbo.TCAJA.lRefresca, dbo.TCAJA.lPasswordC, dbo.TCAJA.lPassword, dbo.TCAJA.tGrupo, dbo.TCAJA.lConsumo1, dbo.TCAJA.lConsumo2, " & _
           " dbo.TCAJA.lConsumo3, dbo.TCAJA.lPrecuenta, dbo.TCAJA.lAdicion, dbo.TCAJA.lPrecuentaAgrupada, dbo.TCAJA.tTipoPedido, dbo.TCAJA.lObliga, dbo.TCAJA.lMozo, dbo.TCAJA.lObligaPrinter, dbo.TCAJA.lPax, dbo.TCAJA.lObligaCierre, dbo.TCAJA.lFiltroTipoPedido, dbo.TCAJA.nPuerto, dbo.TCAJA.tMensaje1, dbo.TCAJA.tMensaje2, " & _
           " dbo.TCAJA.lCancelacion, dbo.TCAJA.lDirecto, dbo.TCAJA.lObligaPrecuenta, dbo.TCAJA.lComboPrecuenta, dbo.TCAJA.lComboDocumento, dbo.TCAJA.lCambioMesa, dbo.TCAJA.lVisaNet, dbo.TCAJA.lImpuestoPrecuenta, dbo.TCAJA.lDocumentoAgrupado, dbo.TCAJA.lOrden, dbo.TCAJA.lValorCortesia, dbo.TCAJA.lObservacion, " & _
           " dbo.TCAJA.lCajaRapida, dbo.TCAJA.lPropiedadPrecuenta, dbo.TCAJA.lPropiedadDocumento, dbo.TCAJA.lPrecioNetoPrecuenta, dbo.TCAJA.nLimitePrecuenta, dbo.TCAJA.tUnidadNegocio, dbo.TCAJA.nLimiteReimpresion, dbo.TCAJA.lPasswordTransferencia, dbo.TCAJA.lCD, dbo.TCAJA.lFechaEntregaDelivery, " & _
           " dbo.TCAJA.lMultiCajero, dbo.TCAJA.lMCPV, dbo.TCAJA.lCCVOX, dbo.TCAJA.lMotorizado, dbo.TCAJA.lEquivaDolaPrecuenta, dbo.TCAJA.tSubAlmacen, dbo.TCAJA.lObservacionPrecuenta, dbo.TCAJA.lObservacionDocumento, dbo.TCAJA.lPasswordImportarPedido, dbo.TCAJA.lActivaImpDscAlternativa, " & _
           " dbo.TCAJA.lCompatibilidadTVS, dbo.TCAJA.nLongitudBarra, dbo.TCAJA.lPagoRapido, dbo.TCAJA.lDisgrega, dbo.TCAJA.lPasswordPorCobrar, dbo.TCAJA.lModificaTipoPedido, dbo.TCAJA.tSucursal, dbo.TCAJA.nBalanzaPuerto, dbo.TCAJA.lCapturaPeso, dbo.TCAJA.lPagoRapidoPV, dbo.TCAJA.tTextoConsumo, " & _
           " dbo.vGrupo.Descripcion AS xGrupo, dbo.TIMPRESORA.tDescripcion AS xImpresora, dbo.vTipoPedido.Descripcion AS xTipoPedido, lSiab, tSectorVenta, lcajamobile, lBloqueaPrecuenta, lRotulado, isnull(lmultiAreaSubGrupo,0) lmultiAreaSubGrupo,isnull(lmultiAreaCaja,0) lmultiAreaCaja, " & _
           " isnull(lHuella,0) lHuella , isnull(lImprimeImagCabPrecuenta,0) lImprimeImagCabPrecuenta , isnull(lImprimeImagpiePrecuenta,0) lImprimeImagpiePrecuenta, isnull(lAccesoDespachoPedido,0) lAccesoDespachoPedido, isnull(lBuscaPedidoNumero,0) lBuscaPedidoNumero, isnull(lCodigoReciboIngreso,0) lCodigoReciboIngreso, " & _
           " isnull(lPagoRapidoMod,0) lPagoRapidoMod, isnull(lWebAp,0) lWebAp,isnull(lMesa247,0) lMesa247, isnull(lConsumo4,0) lConsumo4, isnull(lPrecuentaNoValorizada,0) lPrecuentaNoValorizada, isnull(tcaja.LBuscarpedidovisualizargrilla,0) LBuscarpedidovisualizargrilla, Isnull(tcaja.lbuscarpedidofiltrarmesa,0) lbuscarpedidofiltrarmesa, Isnull(tcaja.lClaveEnvioProduccion,0) lClaveEnvioProduccion, isnull(tcaja.EAN13,0) EAN13 ,  " & _
           " dbo.tcaja.lObservacionCabDoc,isnull(dbo.tcaja.tCompMesa247,'') as tCompMesa247,isnull(dbo.tcaja.lMotivoDescuento,0) as lMotivoDescuento,isnull(dbo.tcaja.lCajaContingencia,0) as lCajaContingencia, isnull(dbo.tcaja.lImpPropina,0) as lImpPropina, isnull(dbo.tcaja.lImpComandaf2,'') as lImpComandaf2, isnull(dbo.tcaja.lPassOtrosPagos,0) as lPassOtrosPagos , isnull(dbo.tcaja.lImpClienteCab,0) as lImpClienteCab, " & _
           " isnull(dbo.tcaja.timpfiscal,'') as tImpFiscal, isnull(dbo.tcaja.lBuscarPedidoBarra,0) as lBuscarPedidoBarra, isnull(dbo.tcaja.lIntVisaNetQR,0) as  lIntVisaNetQR, isnull(lImpComandaBarra,0) as lImpComandaBarra, isnull(lbloqueadocconsumo,0) as lbloqueadocconsumo , isnull(lCashDro,0) as lCashDro, isnull(lIntUber,0) as lIntUber, isnull(tIpCashDrow,'') as tIpCashDrow, isnull(dbo.tcaja.lIntMercadoPagoQR,0) as  lIntMercadoPagoQR, " & _
           " isnull(lIntPatio,0) lIntPatio, isnull(lBloqueoAfectoInafecto,0) as lBloqueoAfectoInafecto, isnull(dbo.tcaja.lIntMercadoPagoQREstatico,0) as  lIntMercadoPagoQREstatico, isnull(dbo.tcaja.lIntPagoEfectivo,0) as  lIntPagoEfectivo , isnull(dbo.tcaja.lPagoDirectoCD,0) as  lPagoDirectoCD , isnull(dbo.tcaja.lImprimePropinaSugerida,0) as lImprimePropinaSugerida  , isnull(dbo.tcaja.lActivarAvisoRetenciones,0) as lActivarAvisoRetenciones,   " & _
           " isnull(dbo.tcaja.lActivaTresDecimales,'') as lActivaTresDecimales, isnull(dbo.tcaja.lIntegracionCASSWII30CD,0) as lIntegracionCASSWII30CD, isnull(tDefaultOrigenVenta,'') tDefaultOrigenVenta, isnull(lNoVisualizarItemCombos,0) lNoVisualizarItemCombos, isnull(lMostrarPrecioItemsCombo,0) lMostrarPrecioItemsCombo, isnull(lActivarMensajeCtsCorrientes,0) lActivarMensajeCtsCorrientes , isnull(lIntMobileKDS,0)lIntMobileKDS, isnull(lImprimirGuiaFormatoCorp,0)lImprimirGuiaFormatoCorp, isnull(lPermitirAnularMaxDias,0)lPermitirAnularMaxDias, isnull(lOrdenarClientes, 0)lOrdenarClientes, lUserToken,lApiKey, lApiToken, isnull(lVistaPreviaItemsCombo,0)lVistaPreviaItemsCombo,  " & _
           " isnull(dbo.tcaja.lImprimePropinaDoc,'') as lImprimePropinaDoc, isnull(dbo.tcaja.IntFPAY,0)IntFPAY , isnull(lActivaCreditoCoorporativo, 0) lActivaCreditoCoorporativo  , isnull(IntegracionWebApi, 0) IntegracionWebApi , isnull(lModoFacturacion, 0) lModoFacturacion , isnull(lActivarImpresionBlu, 0) lActivarImpresionBlu, isnull(lActivaRestriClienteDelivery,0) lActivaRestriClienteDelivery, isnull(dbo.tcaja.lMotivoAdm,0) lMotivoAdm , isnull(lBloqueoAnulaPedido,0)lBloqueoAnulaPedido,  isnull(lBloqueoAnulaItemsPedido,0) lBloqueoAnulaItemsPedido, isnull(lOcultaRazonSocial,0) as lOcultaRazonSocial,isnull(lBloqueaCanal1,0) as lBloqueaCanal1,isnull(lBloqueaCanal2,0) as lBloqueaCanal2, " & _
           " isnull(lBloqueaCanal3,0) as lBloqueaCanal3,isnull(lBloqueaCanal4,0) as lBloqueaCanal4,isnull(lBloqueaCanal5,0) as lBloqueaCanal5, isnull(lImprimePrecuentaCabunica,0) as lImprimePrecuentaCabunica, isnull(lBloquearPedidoSinMesa,0) as lBloquearPedidoSinMesa, isnull(lAutorizacionSinMeza,0) as lAutorizacionSinMeza, isnull(lSolicitaClaveEgreso,0) lSolicitaClaveEgreso, isnull(lBloqueoEdicionPedidoConsolaAgr,0) lBloqueoEdicionPedidoConsolaAgr, isnull(lActivaPasswordMulticajero,0) lActivaPasswordMulticajero, isnull(lImprimePagoPrecuenta,0) lImprimePagoPrecuenta, Isnull(lBusquedaReservas,0) as lBusquedaReservas, Isnull(lOrdenesRappi,0) as lOrdenesRappi, Isnull(tDespachador,'') as tDespachador, Isnull(lActivarValesConsumo,0) as lActivarValesConsumo, Isnull(lActivaIntegracionNiubiz,0) as lActivaIntegracionNiubiz, Isnull(lActivaIntegracionIzipay,0) as lActivaIntegracionIzipay, Isnull(lTurnoCompartido,0) as lTurnoCompartido, Isnull(lAlertaConsola,0) as lAlertaConsola,   " & _
           " Isnull(lActivaVoucher,0) as lActivaVoucher,  " & _
           " Isnull(lActivaAgrupadorCajaRapida,0) as lActivaAgrupadorCajaRapida,  " & _
           " Isnull(tCodigoAgrupadorCajaRapida,'') as tCodigoAgrupadorCajaRapida  " & _
           " FROM dbo.TCAJA LEFT OUTER JOIN dbo.vTipoPedido ON dbo.TCAJA.tTipoPedido = dbo.vTipoPedido.Codigo LEFT OUTER JOIN dbo.TIMPRESORA ON dbo.TCAJA.tCaja = dbo.TIMPRESORA.tCaja AND dbo.TCAJA.tPrecuenta = dbo.TIMPRESORA.tImpresora LEFT OUTER JOIN dbo.vGrupo ON dbo.TCAJA.tGrupo = dbo.vGrupo.Codigo "
           
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(6, grdGrilla, "Codigo", 2, "tCaja", 850, 2, 0, "", _
                                  "Descripción Detallada", 2, "tDescripcion", 2500, 0, 0, "", _
                                  "Canal Predet.", 2, "xTipoPedido", 2000, 0, 0, "", _
                                  "Grupo Predet.", 2, "xGrupo", 2000, 0, 0, "", _
                                  "Imp.PreCuenta", 2, "xImpresora", 2000, 0, 0, "", _
                                  "Activo", 2, "lActivo", 900, 2, 4, "")
        
    Set grdGrilla.DataSource = RsCabecera
   
    cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCaja = Nothing
End Sub

Private Sub grdGrilla_DblClick()
    cmdOpcion_Click (1)
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

