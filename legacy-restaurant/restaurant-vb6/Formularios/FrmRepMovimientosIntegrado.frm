VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmRepMovimientosIntegrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Articulos - Integrado"
   ClientHeight    =   6045
   ClientLeft      =   3360
   ClientTop       =   7350
   ClientWidth     =   10935
   Icon            =   "FrmRepMovimientosIntegrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10935
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
      Left            =   7680
      Picture         =   "FrmRepMovimientosIntegrado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5355
      Width           =   1455
   End
   Begin VB.CommandButton CmdActualiza 
      Caption         =   "Actualizar"
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
      Left            =   6240
      Picture         =   "FrmRepMovimientosIntegrado.frx":03FC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5355
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
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
      Index           =   2
      Left            =   4800
      Picture         =   "FrmRepMovimientosIntegrado.frx":04C4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5355
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
      Left            =   3360
      Picture         =   "FrmRepMovimientosIntegrado.frx":0CE6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5355
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
      Left            =   1920
      Picture         =   "FrmRepMovimientosIntegrado.frx":1218
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5355
      Width           =   1455
   End
   Begin VB.Frame Frame1 
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
      Height          =   5235
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10905
      Begin VB.TextBox TxtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1125
         Width           =   885
      End
      Begin VB.TextBox TxtArticulo 
         BackColor       =   &H80000000&
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
         Left            =   2295
         TabIndex        =   31
         Top             =   1125
         Width           =   2880
      End
      Begin VB.CheckBox ChkArticulo 
         Caption         =   "Todos"
         Height          =   210
         Left            =   5235
         TabIndex        =   30
         Top             =   1200
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.Frame Frame3 
         Caption         =   "Formato"
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   135
         TabIndex        =   27
         Top             =   3630
         Width           =   5985
         Begin VB.OptionButton optFormato 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   1
            Left            =   3345
            TabIndex        =   29
            Top             =   300
            Width           =   2160
         End
         Begin VB.OptionButton optFormato 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   28
            Top             =   300
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Visualización"
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   135
         TabIndex        =   26
         Top             =   2910
         Width           =   5985
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por cantidades"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   35
            Top             =   300
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por montos"
            Height          =   240
            Index           =   1
            Left            =   3360
            TabIndex        =   34
            Top             =   300
            Width           =   2160
         End
      End
      Begin VB.ComboBox cboTipoDocumento 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   3765
      End
      Begin VB.Frame Frame7 
         Caption         =   " Tipo de Reporte "
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   135
         TabIndex        =   17
         Top             =   4380
         Width           =   5985
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por local"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   1
            Left            =   3345
            TabIndex        =   10
            Top             =   285
            Value           =   -1  'True
            Width           =   2160
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por sector"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   0
            Left            =   165
            TabIndex        =   9
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Servidores en línea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5070
         Left            =   6240
         TabIndex        =   16
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar todo"
            Height          =   360
            Left            =   2880
            TabIndex        =   13
            Top             =   4665
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4365
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7699
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.CheckBox ChkFamilia 
         Caption         =   "Todas"
         Height          =   195
         Left            =   5235
         TabIndex        =   6
         Top             =   345
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.CheckBox ChkSubFamilia 
         Caption         =   "Todas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5235
         TabIndex        =   8
         Top             =   750
         Value           =   1  'Checked
         Width           =   810
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   5085
         Top             =   2025
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo CboFamilia 
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   285
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483633
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
      Begin MSDataListLib.DataCombo CboSubFamilia 
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Top             =   705
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483633
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
      Begin MSComCtl2.DTPicker DtpFInicio 
         Height          =   345
         Left            =   1410
         TabIndex        =   20
         Top             =   1980
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51445763
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker DtpFFinal 
         Height          =   345
         Left            =   1410
         TabIndex        =   21
         Top             =   2430
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51445763
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker dtpHInicio 
         Height          =   345
         Left            =   3045
         TabIndex        =   22
         Top             =   1980
         Width           =   1290
         _ExtentX        =   2275
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
         CustomFormat    =   "HH:mm:ss"
         Format          =   51445763
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHFinal 
         Height          =   345
         Left            =   3045
         TabIndex        =   23
         Top             =   2430
         Width           =   1290
         _ExtentX        =   2275
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
         CustomFormat    =   "HH:mm:ss"
         Format          =   51445763
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   4
         Left            =   705
         TabIndex        =   33
         Top             =   1185
         Width           =   600
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Index           =   3
         Left            =   390
         TabIndex        =   25
         Top             =   2490
         Width           =   915
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   24
         Top             =   2055
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Movimiento:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   14
         Top             =   750
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmRepMovimientosIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemporal   As Recordset
Dim rsReporte    As Recordset
Dim RsFamilia    As Recordset
Dim RsSubFamilia As Recordset
Dim RsArea       As Recordset
Dim RsDatos      As Recordset
Dim RsTipoDocumento As Recordset
Dim sNombre      As String

Dim sCriterio, sTituloG, sPrecio, sOrden As String
Dim sFiltro, sEvaluaFiltro, sOrder As String

Dim CrtReporte    As New CrtRepMovimientosIntegrado
Dim CrtReporteDet As New CrtRepMovimientosDetalladoIntegrado

Sub CreaTemporal()
    If OptFormato(0).value Then
        sNombre = dbTemporal(sUsuario, 6, "Tienda", "nVarChar(80)", _
                                          "Codigo", "nVarChar(20)", _
                                          "Descripcion", "nVarChar(200)", _
                                          "Unidad", "nVarChar(50)", _
                                          "Monto", "Float", _
                                          "TipoMov", "nVarchar(30)")
    Else
        sNombre = dbTemporal(sUsuario, 8, "Tienda", "nVarChar(80)", _
                                          "Area", "nVarChar(50)", _
                                          "Codigo", "nVarChar(20)", _
                                          "Descripcion", "nVarChar(300)", _
                                          "Unidad", "nVarChar(50)", _
                                          "Fecha", "nVarChar(50)", _
                                          "Cantidad", "Float", _
                                          "Precio", "Float")
    End If
End Sub

Private Sub configuraListaServidores()
    Dim rsS As New ADODB.Recordset
    Dim Item As ListItem
    Dim X As Integer
    
    Set rsS = devuelveServidoresConectados
    
    With lstServidores
        .Checkboxes = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
        .ColumnHeaders.Add , , " Local ", 1100
        .ColumnHeaders.Add , , " Estado ", 1100
        .ColumnHeaders.Add , , "Ip", 0
        .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
     
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(6)
            rsS.MoveNext
        Next X
    End With
End Sub

Sub LlenaCombo()
    If Not CnAlmacen Is Nothing Then
        Isql = "Select tCodigo as Codigo, tDetallado as Descripcion from tTabla where lActivo='1' and tTabla='FAMILIA' Order By 2"
        Llena_Combo Isql, CboFamilia, RsFamilia, "Codigo", "Descripcion"
        
        Isql = "Select tCodigoSubFamilia as Codigo, tDetallado as Descripcion from tSubFamilia where lActivo='1' Order By 2"
        Llena_Combo Isql, CboSubFamilia, RsSubFamilia, "Codigo", "Descripcion"
        
        Isql = "Select Codigo, Descripcion From vTipoDocumento Where Codigo >=91 Order By Codigo"
        Set RsTipoDocumento = Lib.OpenRecordset(Isql, CnAlmacen)
        i = 1
        If RsTipoDocumento.RecordCount > 0 Then
            cboTipoDocumento.AddItem "Ingresos"
            cboTipoDocumento.ItemData(0) = "00"
            While Not RsTipoDocumento.EOF
                cboTipoDocumento.AddItem RsTipoDocumento!Descripcion
                cboTipoDocumento.ItemData(i) = RsTipoDocumento!Codigo
                i = i + 1
                RsTipoDocumento.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub CboFamilia_Change()
    CboSubFamilia.BoundText = ""
End Sub

Private Sub CboFamilia_LostFocus()
    Isql = "Select tCodigoSubFamilia as Codigo, tDetallado as Descripcion from tSubFamilia where lActivo='1' and tcodigoFamilia='" & CboFamilia.BoundText & "'"
    Llena_Combo Isql, CboSubFamilia, RsSubFamilia, "Codigo", "Descripcion"
End Sub

Private Sub ChkArticulo_Click()
    If ChkArticulo.value = 1 Then
        TxtCodigo.Text = ""
        TxtArticulo = ""
        TxtArticulo.BackColor = &H8000000F
        TxtArticulo.Enabled = False
    Else
        TxtArticulo.Enabled = True
        TxtArticulo.BackColor = &HF0EADB
        TxtArticulo.SetFocus
    End If
End Sub

Private Sub CmdActualiza_Click()
    If MsgBox("Esta operación puede tardar varios minutos" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
        Screen.MousePointer = vbHourglass
        configuraListaServidores
        Me.chkTodo.value = 0
        MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index <> 3 Then
        CreaTemporal
        
        sTituloG = ""
        sOrder = ""
        
        If ChkFamilia.value = 0 Then
            If CboFamilia.BoundText = "" Then
                MsgBox "Seleccione la Familia...", vbCritical, sMensaje
                CboFamilia.SetFocus
                Exit Sub
            End If
            sTituloG = sTituloG & "Familia: " & CboFamilia.Text
        Else
            sTituloG = sTituloG & "Familia: Todas"
        End If
        If ChkSubFamilia.value = 0 Then
            If CboSubFamilia.BoundText = "" Then
                MsgBox "Seleccione la Sub-Familia...", vbCritical, sMensaje
                CboSubFamilia.SetFocus
                Exit Sub
            End If
            sTituloG = sTituloG & " / Sub-Familia: " & CboSubFamilia.Text
        Else
            sTituloG = sTituloG & " / Sub-Familia: Todas"
        End If
        If optVisualizacion(0).value Then
            sTituloG = sTituloG & " / Visualización: Cantidades"
        Else
            sTituloG = sTituloG & " / Visualización: Montos"
        End If
        
        If cboTipoDocumento.ListIndex = -1 Then
            MsgBox "Seleccione el Tipo de Movimiento...", vbCritical, sMensaje
            cboTipoDocumento.SetFocus
            Exit Sub
        End If
        
        For w = 1 To lstServidores.ListItems.Count
            If lstServidores.ListItems.Item(w).Checked = True Then
                lSw = True
                Exit For
            End If
        Next
        If lSw = False Then
            MsgBox "Seleccione una Localidad!", vbCritical, sMensaje
            Exit Sub
        End If
       
        Screen.MousePointer = vbHourglass
        GeneraDatos
        
        If rsReporte.RecordCount = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "No hay datos que mostrar...", vbExclamation, sMensaje
            Set rsReporte = Nothing
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbDefault
    sPrecio = ""
    Select Case Index
        Case Is = 0 'Emite
            frmEmite.CRViewer.DisplayGroupTree = False
            If OptFormato(0).value Then
                CrtReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporte.Database.SetDataSource rsReporte
                CrtReporte.TxtRazonSocial.SetText sRazonComercial
                CrtReporte.Subtitulo.SetText sTituloG
                CrtReporte.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporte.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporte.PaperOrientation = crLandscape
                CrtReporte.DiscardSavedData
                frmEmite.CRViewer.ReportSource = CrtReporte
            Else
                CrtReporteDet.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRazonSocial.SetText sRazonComercial
                CrtReporteDet.Subtitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporteDet.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporteDet.PaperOrientation = crPortrait
                CrtReporteDet.DiscardSavedData
                frmEmite.CRViewer.ReportSource = CrtReporteDet
            End If
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
            If sNombre <> "" Then Cn.Execute "Drop Table " & sNombre
            
        Case Is = 1 'Impresora
            If OptFormato(0).value Then
                CrtReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporte.Database.SetDataSource rsReporte
                CrtReporte.TxtRazonSocial.SetText sRazonComercial
                CrtReporte.Subtitulo.SetText sTituloG
                CrtReporte.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporte.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporte.PaperOrientation = crLandscape
                CrtReporte.DiscardSavedData
                CrtReporte.PrintOut
            Else
                CrtReporteDet.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRazonSocial.SetText sRazonComercial
                CrtReporteDet.Subtitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporteDet.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporteDet.PaperOrientation = crLandscape
                CrtReporteDet.DiscardSavedData
                CrtReporteDet.PrintOut
            End If
            If sNombre <> "" Then Cn.Execute "Drop Table " & sNombre
            
        Case Is = 2 'Exportar
            If OptFormato(0).value Then
                CrtReporte.ExportOptions.FormatType = crEFTExcel50
                CrtReporte.ExportOptions.DestinationType = crEDTDiskFile
                CrtReporte.Database.SetDataSource rsReporte
                CrtReporte.TxtRazonSocial.SetText sRazonComercial
                CrtReporte.Subtitulo.SetText sTituloG
                CrtReporte.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporte.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporte.PaperOrientation = crLandscape
                CrtReporte.DiscardSavedData
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                CrtReporte.ExportOptions.DiskFileName = cmdSave.FileName
                If Trim(CrtReporte.ExportOptions.DiskFileName) <> "" Then CrtReporte.Export False
            Else
                CrtReporteDet.ExportOptions.FormatType = crEFTExcel50
                CrtReporteDet.ExportOptions.DestinationType = crEDTDiskFile
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRazonSocial.SetText sRazonComercial
                CrtReporteDet.Subtitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "Movimientos de Articulos - " & UCase(cboTipoDocumento.Text)
                CrtReporteDet.txtRango.SetText "Del " & DtpFInicio & " " & dtpHInicio & " Al " & DtpFFinal & " " & dtpHFinal
                CrtReporteDet.PaperOrientation = crLandscape
                CrtReporteDet.DiscardSavedData
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                CrtReporteDet.ExportOptions.DiskFileName = cmdSave.FileName
                If Trim(CrtReporteDet.ExportOptions.DiskFileName) <> "" Then CrtReporteDet.Export False
            End If
            If sNombre <> "" Then Cn.Execute "Drop Table " & sNombre
            
        Case Is = 3 'Salir
            Unload Me
    End Select
    
    Screen.MousePointer = vbDefault
    Set rsReporte = Nothing
End Sub

Private Sub ChkFamilia_Click()
    If ChkFamilia.value = 1 Then
       CboFamilia.Enabled = False
       CboFamilia = ""
       CboFamilia.BackColor = &H8000000F
       ChkSubFamilia.value = 1
       ChkSubFamilia.Enabled = False
       CboSubFamilia.Enabled = False
       CboSubFamilia = ""
       CboSubFamilia.BackColor = &H8000000F
       
       ChkArticulo.Enabled = True
       TxtArticulo.Enabled = True
       TxtArticulo.BackColor = &H8000000F
       
    Else
       ChkSubFamilia.Enabled = True
       CboFamilia.Enabled = True
       CboFamilia.BackColor = &HF0EADB
       
       ChkArticulo.Enabled = False
       ChkArticulo.value = 1
       TxtArticulo.Text = ""
       TxtArticulo.Enabled = False
       TxtArticulo.BackColor = &H8000000F
    End If
End Sub

Private Sub ChkSubFamilia_Click()
    If ChkSubFamilia.value = 1 Then
       CboSubFamilia.Enabled = False
       CboSubFamilia = ""
       CboSubFamilia.BackColor = &H8000000F
    Else
       CboSubFamilia.Enabled = True
       CboSubFamilia.BackColor = &HF0EADB
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
   
    Call LlenaCombo
    sEvaluaFiltro = "111"
    DtpFInicio.value = "01/" & str(Month(Date)) & "/" & str(Year(Date))
    DtpFFinal.value = Date
    dtpHInicio.value = "00:00:00"
    dtpHFinal.value = "23:59:59"
    
    obtieneListaServidores
End Sub

Sub GeneraDatos()
    Dim oComando As clsComando
    Set oComando = New clsComando

    Set rsReporte = Nothing
     
    CreaTemporal
    
    '~~~~~~ Pre-Filtro sobre stocks ~~~~~~
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
            conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
            If conectado = False Then
                MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                lstServidores.ListItems.Item(w).Checked = False
                GoTo sigue
            End If
            If OptFormato(0).value Then
                If Not oComando.CreateCmdSp("spRep_MovimientosArticuloIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            Else
                If Not oComando.CreateCmdSp("spRep_MovimientosArticuloDetalladoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            End If
            oComando.CreateParameter "@dat_FechaIni", adVarChar, adParamInput, 20, Format(DtpFInicio, "YYYYmmdd") & " " & Format(dtpHInicio, "HH:mm:ss")
            oComando.CreateParameter "@dat_FechaFin", adVarChar, adParamInput, 20, Format(DtpFFinal, "YYYYmmdd") & " " & Format(dtpHFinal, "HH:mm:ss")
            oComando.CreateParameter "@vch_familia", adVarChar, adParamInput, 20, CboFamilia.BoundText
            oComando.CreateParameter "@vch_subfamilia", adVarChar, adParamInput, 20, CboSubFamilia.BoundText
            oComando.CreateParameter "@vch_Articulo", adVarChar, adParamInput, 20, TxtCodigo.Text
            oComando.CreateParameter "@vch_Visualiza", adVarChar, adParamInput, 1, IIf(optVisualizacion(0).value, "C", "M")
            oComando.CreateParameter "@vch_TipoMov", adVarChar, adParamInput, 2, Format(cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex), "00")
            If Not oComando.GetParamOK Then
                Set oComando = Nothing
                Exit Sub
            End If
            cl = lstServidores.ListItems.Item(w)
            sector = lstServidores.ListItems.Item(w).SubItems(1)
            llocal = lstServidores.ListItems.Item(w).SubItems(2)
            Set RsDatos = oComando.GetSP()
            With RsDatos
                While Not RsDatos.EOF
                    If optOpcion(0).value Then
                        Tienda = sector
                    Else
                        Tienda = llocal
                    End If
                    If OptFormato(0).value Then
                        Isql = "INSERT INTO " & sNombre & "(Tienda, Codigo, Descripcion, Unidad, Monto, TipoMov) " & _
                               "VALUES('" & Tienda & "','" & !Codigo & "', '" & !Descripcion & "','" & !unidad & "'," & !Monto & ",'')"
                    Else
                        Isql = "INSERT INTO " & sNombre & "(Tienda, Area, Codigo, Descripcion, Unidad, Fecha, Cantidad, Precio) " & _
                               "VALUES('" & Tienda & "','" & !Area & "','" & !Codigo & "', '" & !Descripcion & "','" & !unidad & "','" & Format(!Fecha, "dd/MM/yyyy") & "'," & !Cantidad & "," & !Precio & ")"
                    End If
                    Cn.Execute Isql
                    RsDatos.MoveNext
                Wend
            End With
        End If
sigue:
    Next w

    'Realiza el llenado de las tiendas que no tengan asignado
    'For w = 1 To lstServidores.ListItems.Count
    '    If lstServidores.ListItems.Item(w).Checked = True Then
    '        llocal = lstServidores.ListItems.Item(w).SubItems(2)
    '        Isql = "INSERT INTO " & sNombre & " " & _
    '               "SELECT '" & llocal & "', Codigo, Descripcion, Unidad, 0, '' FROM " & sNombre & " GROUP BY Codigo, Descripcion, Unidad HAVING Codigo NOT IN (SELECT Codigo FROM " & sNombre & " WHERE Tienda = '" & llocal & "')"
    '        Cn.Execute Isql
    '    End If
    'Next w
    If OptFormato(0).value Then
        Isql = "SELECT Codigo, upper(Descripcion) AS Descripcion, Unidad, Monto, Tienda FROM " & sNombre
    Else
        Isql = "SELECT Area, Codigo, Descripcion, Unidad, Fecha, Cantidad, Precio, Tienda FROM " & sNombre
    End If
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
End Sub

'integrados
Public Sub activaCheck(Tipo As Boolean)
    Dim o As Integer
    
    For o = 1 To lstServidores.ListItems.Count
        If Tipo = True Then
            If lstServidores.ListItems.Item(o).SubItems(3) <> "Desconectado" Then
                lstServidores.ListItems.Item(o).Checked = Tipo
            End If
        Else
            lstServidores.ListItems.Item(o).Checked = Tipo
        End If
    Next o
End Sub

Private Sub obtieneListaServidores()
    Dim rsS As New ADODB.Recordset
    Dim Item As ListItem
    Dim X As Integer
    
    Set rsS = rsListaServidores
    Me.lstServidores.Checkboxes = True
    lstServidores.ColumnHeaders.Clear
    Me.lstServidores.ListItems.Clear

    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
        .ColumnHeaders.Add , , " Local ", 1100
        .ColumnHeaders.Add , , " Estado ", 1100
        .ColumnHeaders.Add , , "Ip", 0
        .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
    
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1) 'Sector
            Item.SubItems(2) = rsS.Fields(2) 'Local
            Item.SubItems(3) = rsS.Fields(5) 'Estado
            'Item.SubItems(4) = rsS.Fields(3) 'Ip
            Item.SubItems(4) = rsS.Fields(6) 'Ip
            'Item.SubItems(5) = rsS.Fields(6) 'Base de Datos
            Item.SubItems(5) = rsS.Fields(7) 'Base de Datos
        rsS.MoveNext
        Next X
    End With
End Sub

Private Sub chkTodo_Click()
    activaCheck IIf(chkTodo.value = 1, True, False)
End Sub

Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
       End If
End Sub

Private Sub TxtArticulo_Change()
    If Len(Trim(UCase(TxtArticulo))) = 1 Then
        Isql = "Select tProducto.tCodigoProducto As Codigo, tProducto.tDetallado As Descripcion, vUnidadMedida.Descripcion AS UnidadEntrada" & _
               " From tProducto Left Outer Join vUnidadMedida ON tProducto.tUnidadEntrada = vUnidadMedida.Codigo "
        Call ConfGrilla(3, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 2300, 2, 0, "", _
                                                      "Insumo", 2, "Descripcion", 5000, 0, 0, "", _
                                                      "Unidad", 2, "UnidadEntrada", 1000, 0, 0, "")
               
        frmBuscaAlmacen.nPredeterm = 1
        frmBuscaAlmacen.txtCriterio = Trim(UCase(TxtArticulo))
        frmBuscaAlmacen.txtCriterio.SelStart = Len(Trim(frmBuscaAlmacen.txtCriterio))
        sCodigo = ""
        sDescrip = ""
        
        frmBuscaAlmacen.Show vbModal
       
        TxtCodigo.Text = sCodigo
        TxtArticulo.Text = sDescrip
    End If
End Sub
