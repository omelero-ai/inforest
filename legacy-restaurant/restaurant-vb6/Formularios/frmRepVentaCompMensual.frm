VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmRepVentaCompMensual 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta Comparativa Mensual"
   ClientHeight    =   8085
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepVentaCompMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar"
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
      Left            =   3060
      Picture         =   "frmRepVentaCompMensual.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
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
      Index           =   2
      Left            =   4515
      Picture         =   "frmRepVentaCompMensual.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
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
      Left            =   1605
      Picture         =   "frmRepVentaCompMensual.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
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
      Left            =   150
      Picture         =   "frmRepVentaCompMensual.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7410
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6165
      Begin VB.OptionButton OptSel 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3240
         TabIndex        =   30
         Top             =   4680
         Width           =   1095
      End
      Begin VB.OptionButton OptSel 
         Caption         =   "Ninguno"
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
         Index           =   1
         Left            =   4320
         TabIndex        =   29
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CheckBox chkNinguno 
         Caption         =   "Ninguno"
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
         Left            =   4080
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   3960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   3000
         TabIndex        =   21
         Top             =   3000
         Width           =   2295
         Begin VB.OptionButton optOpcion 
            Caption         =   "Comensales"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   23
            Top             =   270
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Pedidos"
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   22
            Top             =   540
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkCFacturados 
         Caption         =   "Valorizar con Cero la Pre Venta Facturada"
         Height          =   405
         Left            =   840
         TabIndex        =   14
         Top             =   4080
         Width           =   4275
      End
      Begin VB.ComboBox CmbMes 
         Height          =   360
         Left            =   3015
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   746
         Width           =   2265
      End
      Begin VB.Frame Frame3 
         Caption         =   "Canales de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2140
         Left            =   495
         TabIndex        =   19
         Top             =   1800
         Width           =   2265
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal5"
            Height          =   240
            Index           =   5
            Left            =   225
            TabIndex        =   11
            Top             =   1530
            Width           =   1950
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal4"
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   10
            Top             =   1230
            Width           =   1950
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal3"
            Height          =   240
            Index           =   3
            Left            =   225
            TabIndex        =   9
            Top             =   915
            Width           =   1950
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal2"
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   8
            Top             =   615
            Width           =   1950
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal1"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   7
            Top             =   315
            Value           =   1  'Checked
            Width           =   1950
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Visualización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   3015
         TabIndex        =   17
         Top             =   1815
         Width           =   2265
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   585
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker dtpAnual 
         Height          =   405
         Left            =   3015
         TabIndex        =   4
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   375
         Left            =   3015
         TabIndex        =   6
         Top             =   1200
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
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
         CustomFormat    =   "HH"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin TrueOleDBGrid80.TDBGrid grdSubGrupos 
         Height          =   2085
         Left            =   240
         TabIndex        =   27
         Top             =   5160
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3678
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   2
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   12632256
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
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DeadAreaBackColor=   16777215
         ScrollTrack     =   -1  'True
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.fgcolor=&H8B1512&"
         _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.bgcolor=&H8CFFFF&"
         _StyleDefs(22)  =   "Splits(0).Style:id=67,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=88,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=69,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=71,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=87,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=89,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=90,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=94,.parent=67"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=68,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=69"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=71"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccionar Subgrupo :"
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
         Left            =   840
         TabIndex        =   28
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
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
         Index           =   2
         Left            =   2370
         TabIndex        =   20
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Cierre y Apertura  :"
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
         Index           =   1
         Left            =   495
         TabIndex        =   18
         Top             =   1275
         Width           =   2355
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
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
         Index           =   0
         Left            =   2385
         TabIndex        =   16
         Top             =   315
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   600
      Top             =   6765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Valorización :"
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
      Index           =   3
      Left            =   0
      TabIndex        =   24
      Top             =   1920
      Width           =   1170
   End
End
Attribute VB_Name = "frmRepVentaCompMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDetalleM As String
Dim sDetalleMC As String
Dim ReporteMensual As New dsrVentaCompMensualDia
Dim RsTemporal As Recordset
Dim RsLocal As Recordset
Dim rsReporte As Recordset
Dim meses(11) As String
Dim sPrecio As String
Dim sTexto As String
'pierre
Dim VentMens_RsSubGrupos As Recordset
Dim sSubGrupos As String
Dim colu As TrueOleDBGrid80.Columns
Dim Criterio As String
Dim Hay_Datos As Integer

Sub LlenaCombos()
Dim i As Integer
   For i = 0 To 11
       CmbMes.AddItem meses(i)
   Next i
   CmbMes.ListIndex = Month(Date) - 1
End Sub

Private Sub chkCanal_Click(Index As Integer)
   Dim i As Integer
   i = i + IIf(chkCanal(1).Value, 1, 0)
   i = i + IIf(chkCanal(2).Value, 1, 0)
   i = i + IIf(chkCanal(3).Value, 1, 0)
   i = i + IIf(chkCanal(4).Value, 1, 0)
   i = i + IIf(chkCanal(5).Value, 1, 0)
   If i > 3 Then
      chkCanal(Index).Value = False
      MsgBox "Puede comparar máximo 3 canales de venta", vbInformation, sMensaje
   End If
End Sub

Sub Validar_Todos_Ninguno()
'Cn.Execute "select * into popopo2 from " & sSubGrupos
Dim RS0, RS1, RS2 As Recordset
Set RS0 = Lib.OpenRecordset("Select * from " & sSubGrupos, Cn)
'Set RS1 = Lib.OpenRecordset("Select * from " & sSubGrupos & " where filtrar=0", Cn) 'No Todos
Set RS1 = VentMens_RsSubGrupos.Clone
'Set RS2 = Lib.OpenRecordset("Select * from " & sSubGrupos & " where filtrar=-1", Cn) 'No Ninguno
Set RS2 = VentMens_RsSubGrupos.Clone

RS1.Filter = "Filtrar=0"
RS2.Filter = "Filtrar=-1"

'If RS1.RecordCount = 0 Then
'OptSel(0).value = True
'Else
'OptSel(0).value = False
'End If
'
'If RS2.RecordCount = 0 Then
'OptSel(1).value = True
'Else
'OptSel(1).value = False
'End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
      
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If

'                Cn.Execute " select * into xx1 from " & sSubGrupos
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If Hay_Datos = 0 Then
               Screen.MousePointer = vbDefault
                Exit Sub
                End If
                
               If rsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               frmEmite.CRViewer.DisplayGroupTree = False
               ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteMensual.PaperOrientation = crLandscape
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
                    
          Case Is = 1 ' Imprimir
               Genera
               If Hay_Datos = 0 Then
               Screen.MousePointer = vbDefault
                Exit Sub
                End If
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteMensual.PaperOrientation = crLandscape
               ReporteMensual.PrintOut

          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               If Hay_Datos = 0 Then
               Screen.MousePointer = vbDefault
                Exit Sub
                End If
                Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                
               Screen.MousePointer = vbDefault
               ReporteMensual.ExportOptions.FormatType = 21
               ReporteMensual.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               ReporteMensual.ExportOptions.DiskFileName = cmdSave.FileName
               ReporteMensual.Export False
              
   End Select
End Sub

Private Sub Form_Load()

     sSubGrupos = dbTemporal(sCaja, 5, "Filtrar", "bit", "Codigo", "nVarchar(4)", _
                                     "Descripcion", "nVarchar(50)", "Fila", "nVarchar(4)", _
                                     "Grupo", "nVarchar(50)")
    
    
    
    Cn.Execute "insert into " & sSubGrupos & "(Filtrar, codigo, descripcion, fila, grupo) select convert(bit,'1'), Codigo, Descripcion, ROW_NUMBER() OVER(ORDER BY Descripcion asc), Grupo from vsubgrupo order by descripcion"

    
    
    Set VentMens_RsSubGrupos = Lib.OpenRecordset(" select Filtrar, Codigo, Descripcion, Fila, Grupo from " & sSubGrupos & " order by descripcion", Cn)

     Call ConfGrilla(5, grdSubGrupos, "Filtrar", 2, "Filtrar", 500, 2, 4, "", _
                                    "Codigo", 2, "Codigo", 0, 0, 0, "", _
                                   "Descripcion", 2, "Descripcion", 3000, 0, 0, "", _
                                   "Fila", 2, "Fila", 0, 0, 0, "", _
                                   "Grupo", 2, "Grupo", 1500, 0, 0, "")
    
    
    grdSubGrupos.DataSource = VentMens_RsSubGrupos
   OptSel(0).Value = True
    
    With grdSubGrupos
        .Columns(1).Visible = False
        .Columns(3).Visible = False
    End With
    
    Centrar Me
    dtpAnual.Value = Date

    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "setiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    
    chkCanal(1).Caption = IIf(Trim(sBoton1) = "", "no activo", sBoton1)
    chkCanal(2).Caption = IIf(Trim(sBoton2) = "", "no activo", sBoton2)
    chkCanal(3).Caption = IIf(Trim(sBoton3) = "", "no activo", sBoton3)
    chkCanal(4).Caption = IIf(Trim(sBoton4) = "", "no activo", sBoton4)
    chkCanal(5).Caption = IIf(Trim(sBoton5) = "", "no activo", sBoton5)
    If sBoton1 = "" Then chkCanal(1).Enabled = False
    If sBoton2 = "" Then chkCanal(2).Enabled = False
    If sBoton3 = "" Then chkCanal(3).Enabled = False
    If sBoton4 = "" Then chkCanal(4).Enabled = False
    If sBoton5 = "" Then chkCanal(5).Enabled = False
    LlenaCombos
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set RsTemporal = Nothing
   Set rsReporte = Nothing
   Set VentMens_RsSubGrupos = Nothing
    Cn.Execute "drop table " & sSubGrupos
End Sub

Public Sub Genera()
    Screen.MousePointer = vbHourglass
    Dim oComando As clsComando
    Dim sFecha1 As String
    Dim sFecha2 As String
    Dim sAno1 As String
    Dim sAno2 As String
    Dim sMes As String
    Dim sDia As String
    Dim sTipo(3) As String
    Dim j, i As Integer
    Set oComando = New clsComando
    
    j = 0
    
    For i = 1 To 5
        If chkCanal(i).Value Then
           j = j + 1
           sTipo(j) = "0" & LTrim(str(i))
           If j = 1 Then
              ReporteMensual.Text4.SetText chkCanal(i).Caption
           End If
           If j = 2 Then
              ReporteMensual.Text59.SetText chkCanal(i).Caption
           End If
           If j = 3 Then
              ReporteMensual.Text61.SetText chkCanal(i).Caption
           End If
        End If
    Next i
        
        If chkCFacturados.Value = 0 Then
                If optValor(0).Value Then
                   sPrecio = "dbo.DPEDIDO.nVenta"
                   sTexto = " con Precios de Venta"
                Else
                   sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
                   sTexto = " con Precios Netos"
                End If
        Else
                If optValor(0).Value Then
                        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nVenta end "
                        sTexto = " con Precios de Venta"
                    Else
                        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
                        sTexto = " con Precios Netos"
                End If
        End If
    sFecha1 = ""
    sFecha2 = ""
    sAno1 = LTrim(str(Year(dtpAnual.Value)))
    sAno2 = LTrim(str(Year(dtpAnual.Value) - 1))
    sMes = IIf(CmbMes.ListIndex + 1 > 9, LTrim(str(CmbMes.ListIndex + 1)), "0" & LTrim(str(CmbMes.ListIndex + 1)))
        
    Select Case CmbMes.ListIndex + 1
    
           Case 2
                If Year(dtpAnual.Value) Mod 4 = 0 Then
                   sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/29')"
                   sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/29')"
                Else
                   sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/28')"
                   sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/28')"
                End If
                
           Case 1, 3, 5, 7, 8, 10, 12
                   sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/31')"
                   sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/31')"
                   
           Case 4, 6, 9, 11
                   sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/" & sMes & "/30')"
                   sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/" & sMes & "/30')"
    End Select
    
    If OptSel(0).Value = True Then
     Criterio = ""
     Hay_Datos = 1
     ElseIf OptSel(1).Value = True Then
        Dim m As Integer
        Dim o As Integer
        Dim cadena As String
        cadena = "'"
        o = 0
        j = 0
        For m = 1 To grdSubGrupos.ApproxCount
            If grdSubGrupos.Columns(0).CellValue(m) = True Then
                o = o + 1
                If Len(cadena) = 1 Then
                    cadena = cadena + grdSubGrupos.Columns(1).CellValue(m) + "'"
                Else
                    cadena = cadena + ",'" + grdSubGrupos.Columns(1).CellValue(m) + "'"
                End If
            End If
        Next m
        If o = 0 Then
            MsgBox "Debe escoger minimo un subgrupo", vbCritical, sMensaje
            Hay_Datos = 0
            Exit Sub
        End If
 'MsgBox grdSubGrupos.Columns(1).CellValue(1)
'    ElseIf OptSel(0).value = False And OptSel(1).value = False Then
  
        Criterio = " and tCodigoSubGrupo in (" + cadena + ")" 'select codigo from " & sSubGrupos & " where filtrar=1) "
     '   Cn.Execute "select * into popo1 from " & sSubGrupos
        Hay_Datos = 1
    End If

    If Not oComando.CreateCmdSp("spRep_VentaCompMensual", Cn) Then
      Set oComando = Nothing
      Exit Sub
     End If
     
     oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 100, sPrecio
     oComando.CreateParameter "@sAno1", adVarChar, adParamInput, 4, sAno1
     oComando.CreateParameter "@sAno2", adVarChar, adParamInput, 4, sAno2
     oComando.CreateParameter "@sMes", adVarChar, adParamInput, 4, sMes
     oComando.CreateParameter "@sTipo1", adVarChar, adParamInput, 4, sTipo(1)
     oComando.CreateParameter "@sTipo2", adVarChar, adParamInput, 4, sTipo(2)
     oComando.CreateParameter "@sTipo3", adVarChar, adParamInput, 4, sTipo(3)
     oComando.CreateParameter "@sFecha1", adVarChar, adParamInput, 400, sFecha1
     oComando.CreateParameter "@sFecha2", adVarChar, adParamInput, 400, sFecha2
     oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
     oComando.CreateParameter "@Opcion", adVarChar, adParamInput, 10, IIf(Me.optOpcion(0).Value = True, "PAX", "VENTA")
     oComando.CreateParameter "@criterio", adVarChar, adParamInput, 500, Criterio
     

   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If
      Set rsReporte = oComando.GetSP()

    ReporteMensual.DiscardSavedData
    ReporteMensual.Database.SetDataSource rsReporte
    ReporteMensual.PaperOrientation = crLandscape
    ReporteMensual.ReportTitle = " Comparativo del  Mes de : " & CmbMes.Text & sTexto
    ReporteMensual.anio1VS.SetText dtpAnual.Year
    ReporteMensual.Anio2VS.SetText dtpAnual.Year - 1
    ReporteMensual.anio1PS.SetText dtpAnual.Year
    ReporteMensual.anio2PS.SetText dtpAnual.Year - 1
    ReporteMensual.anio1VD.SetText dtpAnual.Year
    ReporteMensual.anio2VD.SetText dtpAnual.Year - 1
    ReporteMensual.anio1PD.SetText dtpAnual.Year
    ReporteMensual.anio2PD.SetText dtpAnual.Year - 1
    ReporteMensual.anio1VL.SetText dtpAnual.Year
    ReporteMensual.anio2VL.SetText dtpAnual.Year - 1
    ReporteMensual.anio1PL.SetText dtpAnual.Year
    ReporteMensual.anio2PL.SetText dtpAnual.Year - 1
    ReporteMensual.AnioD1.SetText dtpAnual.Year
    ReporteMensual.AnioD2.SetText dtpAnual.Year - 1
    ReporteMensual.Text20.SetText sRazonSocial
    ReporteMensual.Text9.SetText localConectado
    ReporteMensual.Text7.SetText "Total Año " & dtpAnual.Year & " :"
    ReporteMensual.Text6.SetText "Total Año " & dtpAnual.Year - 1 & " :"
    
    If Me.optOpcion(0).Value = True Then
       ReporteMensual.Text22.SetText "Comensales"
       ReporteMensual.Text27.SetText "Comensales"
       ReporteMensual.Text31.SetText "Comensales"
    Else
       ReporteMensual.Text22.SetText "Pedidos"
       ReporteMensual.Text27.SetText "Pedidos"
       ReporteMensual.Text31.SetText "Pedidos"
    End If
    
    
    frmEmite.CRViewer.ReportSource = ReporteMensual
End Sub

Private Sub grdSubGrupos_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 0 Then

Validar_Todos_Ninguno
End If
End Sub

Private Sub grdSubGrupos_FilterChange()
On Error GoTo errHandler
    Set colu = grdSubGrupos.Columns
    Dim c As Integer
    c = grdSubGrupos.col
    grdSubGrupos.HoldFields
    VentMens_RsSubGrupos.Requery
    VentMens_RsSubGrupos.Filter = getFilter(colu)
    grdSubGrupos.col = c
    grdSubGrupos.EditActive = True
    Exit Sub

errHandler:
    'MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(colu)
    VentMens_RsSubGrupos.Filter = adFilterNone
End Sub

'Private Sub grdSubGrupos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'    Dim columna, fila As Integer
'    Dim bookMark As Variant
'    Dim codigo As String
'    Dim falso As Boolean
'    columna = grdSubGrupos.ColContaining(X)
'    fila = grdSubGrupos.RowContaining(Y)
'    If fila >= 0 Then
'    bookMark = grdSubGrupos.RowBookmark(fila)
'
'    If bookMark <> "" Then
'        codigo = grdSubGrupos.Columns(1).CellValue(bookMark)
'        Dim filtro As String
'            If codigo <> "" Then
'            filtro = Calcular("select filtrar as codigo from " & sSubGrupos & " where codigo = '" & codigo & "'", Cn)
'            If filtro = "Verdadero" Then
'            Cn.Execute "update " & sSubGrupos & " set Filtrar=0 where codigo ='" & codigo & "' and Filtrar=1"
'            ElseIf filtro = "Falso" Then
'            Cn.Execute "update " & sSubGrupos & " set Filtrar=1 where codigo ='" & codigo & "' and Filtrar=0"
'            End If
'
'        End If
'        '    TDBGrid1.DataSource = Nothing
'        '    grdSubGrupos.DataSource = Nothing
'        'VentMens_RsSubGrupos.Requery
'        Set VentMens_RsSubGrupos = Lib.OpenRecordset("Select * from " & sSubGrupos & " order by descripcion", Cn)
'
'        '    TDBGrid1.DataSource = VentMens_RsSubGrupos
'        '    grdSubGrupos.DataSource = VentMens_RsSubGrupos
'
'        grdSubGrupos.DataSource = VentMens_RsSubGrupos
'        '    TDBGrid1.Refresh
'        '    grdSubGrupos.Refresh
'        Validar_Todos_Ninguno
'        'grdSubGrupos.col = columna
'        grdSubGrupos.Row = fila
'        grdSubGrupos.col = columna
'        grdSubGrupos.SetFocus
'
'        Set colu = grdSubGrupos.Columns
'        Call cmdClearFilter_Click(colu)
'        VentMens_RsSubGrupos.Filter = adFilterNone
'        End If
'    Else
'    Exit Sub
'    End If
'End If
'End Sub

Private Sub OptSel_Click(Index As Integer)
If Index = 0 Then
    Cn.Execute "update " & sSubGrupos & " set Filtrar=-1"
    
ElseIf Index = 1 Then
    Cn.Execute "update " & sSubGrupos & " set Filtrar=0"
End If

   'VentMens_RsSubGrupos.Requery
   Set VentMens_RsSubGrupos = Lib.OpenRecordset(" select Filtrar, Codigo, Descripcion, Fila, Grupo from " & sSubGrupos & " order by descripcion", Cn)
   VentMens_RsSubGrupos.Filter = adFilterNone
   grdSubGrupos.DataSource = VentMens_RsSubGrupos
End Sub
