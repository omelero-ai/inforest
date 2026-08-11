VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmOrdenesConsola 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8400
   ClientLeft      =   2535
   ClientTop       =   1725
   ClientWidth     =   14160
   Icon            =   "frmOrdenesConsola.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   14160
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   14040
      Top             =   2040
   End
   Begin VB.PictureBox ImageBarra 
      Height          =   495
      Left            =   13920
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraGrilla 
      Height          =   7620
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13875
      Begin VB.Frame PanelMensaje 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   4560
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Imprimiendo QR del Pedio Nº 2500012173"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   600
            Width           =   3975
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7365
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   13680
         _ExtentX        =   24130
         _ExtentY        =   12991
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
         Left            =   450
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   13920
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   14100
      TabIndex        =   0
      Top             =   7650
      Width           =   14160
      Begin VB.CommandButton btnEnviar 
         Caption         =   "Enviar"
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
         Left            =   5880
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
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
         Left            =   8520
         Picture         =   "frmOrdenesConsola.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   50
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5085
         TabIndex        =   1
         Top             =   30
         Width           =   5145
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3450
            Picture         =   "frmOrdenesConsola.frx":0534
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmOrdenesConsola.frx":0A76
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4530
            Picture         =   "frmOrdenesConsola.frx":0FB8
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
            Picture         =   "frmOrdenesConsola.frx":14FA
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
            Picture         =   "frmOrdenesConsola.frx":1A3C
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
            Picture         =   "frmOrdenesConsola.frx":1F7E
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
            Width           =   1740
         End
      End
   End
End
Attribute VB_Name = "frmOrdenesConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim Rsss As Recordset
Dim RsReporte As Recordset
Dim RsCodigoPCD As ADODB.Recordset
Dim nColumna As Integer
Dim xPedido As String
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean
Dim StoreId As String
Dim OrderIdd As String
Dim IdLocall As Integer
Dim CodPedido As String

Private Sub btnEnviar_Click()
    
    If RsCabecera.RecordCount = 0 Then
        Exit Sub
    End If
    
    If (Calcular("Select count(*) as codigo From OrdenesRappiDespachador Where Aplicativo = 'RAPPI' and tCodigoPedido = '" & RsCabecera!tCodigoPedido & "' and tDocumento = '" & RsCabecera!tDocumento & "' and OrderId = '" & RsCabecera!OrderId & "'", Cn)) = 0 Then
        MsgBox "Primero debe generar el Qr para que se pueda enviar los datos de bolsas y bebidas.", vbExclamation, sMensaje
        Exit Sub
    End If

    Dim PedidoCodigo As String
    
    frmEnvioDatosRappi.StoreId = RsCabecera!CodigoIntegracion
    frmEnvioDatosRappi.OrderIdd = RsCabecera!OrderId
    frmEnvioDatosRappi.IdLocall = RsCabecera!IdLocal
    frmEnvioDatosRappi.CodPedido = RsCabecera!tCodigoPedido
    
    frmEnvioDatosRappi.Show vbModal
    
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

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me

    ' Ingrese el Titulo
    Me.Caption = " Ordenes Rappi "
    grdGrilla.Caption = Me.Caption
    nColumna = 2

    Set RsCabecera = Lib.OpenRecordset("exec SP_DepachoPedidosRappi " & 1 & "", Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(10, grdGrilla, "Pedido", 2, "tCodigoPedido", 2000, 2, 0, "", _
                                  "OrderId", 2, "OrderId", 2000, 2, 0, "", _
                                  "Documento", 2, "tDocumento", 2200, 2, 0, "", _
                                  "Aplicativo", 2, "Aplicativo", 1800, 2, 0, "", _
                                  "Total", 2, "Total", 1500, 2, 0, "", _
                                  "Tienda", 2, "Tienda", 3000, 2, 0, "", _
                                  "Fecha", 2, "FechaCreacion", 1800, 2, 0, "dd/MM HH:mm", _
                                  "N° Bolsas", 2, "Cantidad", 1500, 2, 0, "", _
                                  "Bebidas", 2, "lBebida", 1000, 2, 4, "", _
                                  "QR", 2, "GenerarQr", 1800, 2, 0, "")
                                   

    grdGrilla.Columns(9).DividerStyle = dbgNoDividers
    grdGrilla.Columns(9).Font.Underline = True

    Set grdGrilla.DataSource = RsCabecera
    Screen.MousePointer = vbDefault
    
    Timer2.Enabled = False
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set frmDespachador = Nothing
End Sub
Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Generar Qr
    If grdGrilla.col = 9 Then
    
        If RsCabecera.RecordCount > 0 Then
        
            If Len(grdGrilla.Columns(2).Text) > 0 Then
                
                RsCabecera.Find ("OrderId='" & grdGrilla.Columns(1).Text & "'")
                
                StoreId = RsCabecera.Fields("CodigoIntegracion")
                OrderIdd = RsCabecera.Fields("OrderId")
                IdLocall = RsCabecera.Fields("IdLocal")
                CodPedido = RsCabecera.Fields("tCodigoPedido")
                               
                cmdDetalle_Click (1)
                grdGrilla.col = 0
                
            Else
                MsgBox "Por favor primero debe generar el documento para que pueda generar el Qr de esta Orden.", vbExclamation, sMensaje
            End If
        Else
            MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
        End If
    End If
End Sub
Private Sub Timer1_Timer()
On Error GoTo fin
    If RsCabecera.RecordCount > 0 And RsCabecera.AbsolutePosition > 0 Then
        Dim s As Integer
        s = RsCabecera.AbsolutePosition
        RsCabecera.Requery
        RsCabecera.AbsolutePosition = s
    Else
        RsCabecera.Requery
    End If
    Exit Sub
fin:
    RsCabecera.Requery
End Sub
Private Sub cmdDetalle_Click(Index As Integer)
On Error GoTo err1
   Select Case Index
        Case Is = 1 'Imprimir Qr
        
                MostrarMensaje
                                            
                Timer2.Enabled = True
   End Select
Exit Sub
err1:
     MsgBox "Error: " & vbCrLf & Err.Description, vbCritical + vbOKOnly, sMensaje
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub MostrarMensaje()
    PanelMensaje.Left = (Me.Width - PanelMensaje.Width) / 2
    PanelMensaje.Top = (Me.Height - PanelMensaje.Height) / 2
    Label1.Caption = "Imprimiendo QR del Pedio Nº " & CodPedido
    PanelMensaje.Visible = True
End Sub

Private Sub OcultarMensaje()
    PanelMensaje.Visible = False
End Sub
Private Sub Timer2_Timer()
On Error GoTo fin2
    Timer2.Enabled = False
    ObtenerQr
    OcultarMensaje
fin2:
    OcultarMensaje
End Sub
Public Sub ObtenerQr()
    Dim sql As String
    Dim Valor As String
    Dim Rsss As ADODB.Recordset
    
'    Dim CodigoIntegracion As String, OrderId As String
'    Dim IdLocal As Integer
    
    If (Calcular("Select count(*) as codigo From OrdenesRappiDespachador Where Aplicativo = 'RAPPI' and StoreId = '" & StoreId & "' and OrderId = '" & OrderIdd & "'", Cn)) > 0 Then
        Valor = Calcular("Select CodProductoConfirmado as Codigo From OrdenesRappiDespachador Where Aplicativo = 'RAPPI' and StoreId = '" & StoreId & "' and OrderId = '" & OrderIdd & "'", Cn)
    Else
        sql = " exec  SP_DepachoPedidosRappi " & 2 & ", '" & StoreId & "', '" & OrderIdd & "'," & IdLocall & ", '','', 0,false "

        Set Rsss = Lib.OpenRecordset(sql, Cn)

        If Rsss.EOF = False Then
            If Rsss.RecordCount > 0 Then
                Rsss.MoveFirst
                Valor = Rsss!respuesta
            End If
        End If
    End If

    If Valor = "" Then
        MsgBox "Ocurrió un error al generar el código QR. Por favor, revise los registros de errores del sistema para obtener más detalles.", vbExclamation, sMensaje
    Else
        ImageBarra.Picture = LoadPicture(CrearImagenQR_Comanda(Valor))

        ImprimirQR RsCabecera, sDespachador, ImageBarra
    End If
End Sub

Public Sub ImprimirQR(RsImpresion As Recordset, Impresora As String, Optional ImagenBarra As PictureBox = 0)

    Printer.FontBold = True
    
    If Not Imprimir(Impresora) Then
       Exit Sub
    End If
    
    Printer.FontName = "Arial"
    
    Dim tamanio As Integer
    
    tamanio = Printer.FontSize

    Printer.FontSize = tamanio

    Printer.FontName = sFont
    Printer.FontSize = 10
    Printer.FontBold = False
    
     'Cabecera
    Printer.Print "Pedido         : " & RsImpresion!tCodigoPedido
    Printer.Print "Documento : " & RsImpresion!tDocumento
    Printer.Print "Tipo             : " & RsImpresion!TipoPedido
    Printer.Print "Caja             : " & sCaja & IIf(IsNull(sUsuario) Or sUsuario = "", "", IIf(Mid(sUsuario, 1, 1) = "*", " - " & Mid(sUsuario, 2, 15), " - " & sUsuario))
    ImprimeXLinea "Cliente         : " & IIf(Mid(Trim(RsImpresion!Cliente), 1, 1) = "-", Mid(Trim(RsImpresion!Cliente), 2, 15), Trim(RsImpresion!Cliente)), 40, 13
    
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    
    Printer.PaintPicture ImagenBarra, 1000, Printer.CurrentY, 1500, 1500
    
    Printer.FontName = "Arial"

    Printer.FontName = sFont
    Printer.FontSize = tamanio
    Printer.FontBold = False
    
    Printer.EndDoc
    
End Sub
