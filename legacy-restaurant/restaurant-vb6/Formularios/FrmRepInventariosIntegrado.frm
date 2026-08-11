VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmRepInventariosIntegrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Valorizado al Día - Integrado"
   ClientHeight    =   5970
   ClientLeft      =   3105
   ClientTop       =   2625
   ClientWidth     =   10935
   Icon            =   "FrmRepInventariosIntegrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
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
      Left            =   7590
      Picture         =   "FrmRepInventariosIntegrado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
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
      Left            =   6150
      Picture         =   "FrmRepInventariosIntegrado.frx":03FC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
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
      Left            =   4710
      Picture         =   "FrmRepInventariosIntegrado.frx":04C4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
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
      Left            =   3270
      Picture         =   "FrmRepInventariosIntegrado.frx":0CE6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
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
      Left            =   1830
      Picture         =   "FrmRepInventariosIntegrado.frx":1218
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
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
      Height          =   5175
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10920
      Begin VB.Frame Frame4 
         Caption         =   "Area a Seleccionar"
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   135
         TabIndex        =   29
         Top             =   1680
         Width           =   5985
         Begin VB.OptionButton optarea 
            Caption         =   "Almacen central"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   345
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optarea 
            Caption         =   "Sub-almacenes"
            Height          =   240
            Index           =   1
            Left            =   3360
            TabIndex        =   30
            Top             =   345
            Width           =   2160
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Visualización"
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   135
         TabIndex        =   26
         Top             =   2550
         Width           =   5985
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por montos"
            Height          =   240
            Index           =   1
            Left            =   3360
            TabIndex        =   28
            Top             =   330
            Width           =   2160
         End
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por cantidades"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   330
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.CheckBox ChkSubFamilia 
         Caption         =   "Todas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5040
         TabIndex        =   20
         Top             =   855
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.CheckBox ChkFamilia 
         Caption         =   "Todas"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   450
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.CheckBox ChkArticulo 
         Caption         =   "Todos"
         Height          =   210
         Left            =   5040
         TabIndex        =   18
         Top             =   1305
         Value           =   1  'Checked
         Width           =   900
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
         Left            =   2100
         TabIndex        =   17
         Top             =   1230
         Width           =   2880
      End
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
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1230
         Width           =   885
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   5100
         Top             =   4290
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame7 
         Caption         =   " Tipo de Reporte "
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   135
         TabIndex        =   15
         Top             =   4290
         Width           =   5985
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   1
            Left            =   3375
            TabIndex        =   9
            Top             =   345
            Value           =   -1  'True
            Width           =   1980
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Top             =   345
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Servidores en linea"
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
         TabIndex        =   14
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar todo"
            Height          =   360
            Left            =   2880
            TabIndex        =   12
            Top             =   4665
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4365
            Left            =   120
            TabIndex        =   10
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
      Begin VB.Frame Frame3 
         Caption         =   " Incluir artículos con Stock "
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   135
         TabIndex        =   13
         Top             =   3420
         Width           =   5985
         Begin VB.CheckBox ChkFiltro 
            Caption         =   "Negativo"
            Height          =   285
            Index           =   0
            Left            =   315
            TabIndex        =   5
            Top             =   330
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.CheckBox ChkFiltro 
            Caption         =   "En Cero"
            Height          =   285
            Index           =   1
            Left            =   2205
            TabIndex        =   6
            Top             =   330
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.CheckBox ChkFiltro 
            Caption         =   "Positivo"
            Height          =   285
            Index           =   2
            Left            =   4320
            TabIndex        =   7
            Top             =   330
            Value           =   1  'Checked
            Width           =   990
         End
      End
      Begin MSDataListLib.DataCombo CboFamilia 
         Height          =   315
         Left            =   1215
         TabIndex        =   21
         Top             =   390
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
         Left            =   1215
         TabIndex        =   22
         Top             =   810
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   25
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   585
         TabIndex        =   24
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   4
         Left            =   510
         TabIndex        =   23
         Top             =   1290
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmRepInventariosIntegrado"
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
Dim sNombre      As String

Dim sCriterio, sTituloG, sPrecio, sOrden As String
Dim sFiltro, sEvaluaFiltro, sOrder As String

Dim CrtReporte    As New CrtRepInventarioIntegrado

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

Private Sub ChkFiltro_Click(Index As Integer)
    sEvaluaFiltro = Trim(str(ChkFiltro(0))) + Trim(str(ChkFiltro(1))) + Trim(str(ChkFiltro(2)))
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
        sTituloG = ""

        Cn.Execute "Delete From " & sNombre
        
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
    Select Case Index
        Case Is = 0 'Emite
            frmEmite.CRViewer.DisplayGroupTree = False
            CrtReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            CrtReporte.Database.SetDataSource rsReporte
            CrtReporte.TxtRazonSocial.SetText sRazonComercial
            CrtReporte.Subtitulo.SetText sTituloG
            CrtReporte.txtArea.SetText IIf(optarea(0).value, "Area : Almacen Central", "Area : SubAlmacenes")
            CrtReporte.PaperOrientation = crLandscape
            CrtReporte.DiscardSavedData
            frmEmite.CRViewer.ReportSource = CrtReporte
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
            
        Case Is = 1 'Impresora
            CrtReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            CrtReporte.Database.SetDataSource rsReporte
            CrtReporte.TxtRazonSocial.SetText sRazonComercial
            CrtReporte.Subtitulo.SetText sTituloG
            CrtReporte.txtArea.SetText IIf(optarea(0).value, "Area : Almacen Central", "Area : SubAlmacenes")
            CrtReporte.PaperOrientation = crLandscape
            CrtReporte.DiscardSavedData
            CrtReporte.PrintOut
            
        Case Is = 2 'Exportar
            CrtReporte.ExportOptions.FormatType = crEFTExcel50
            CrtReporte.ExportOptions.DestinationType = crEDTDiskFile
            CrtReporte.Database.SetDataSource rsReporte
            CrtReporte.TxtRazonSocial.SetText sRazonComercial
            CrtReporte.Subtitulo.SetText sTituloG
            CrtReporte.txtArea.SetText IIf(optarea(0).value, "Area : Almacen Central", "Area : SubAlmacenes")
            CrtReporte.PaperOrientation = crLandscape
            CrtReporte.DiscardSavedData
            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
            cmdSave.ShowSave
            CrtReporte.ExportOptions.DiskFileName = cmdSave.FileName
            If Trim(CrtReporte.ExportOptions.DiskFileName) <> "" Then CrtReporte.Export False
            
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
    
    sNombre = dbTemporal(sUsuario, 5, "Tienda", "nVarChar(80)", _
                                      "Codigo", "nVarChar(20)", _
                                      "Descripcion", "nVarChar(200)", _
                                      "Monto", "Float", _
                                      "Unidad", "nVarChar(20)")
    obtieneListaServidores
End Sub

Sub GeneraDatos()
    Dim oComando As clsComando
    Set oComando = New clsComando

    Set rsReporte = Nothing
     
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
            conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
            If conectado = False Then
                MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                lstServidores.ListItems.Item(w).Checked = False
                GoTo sigue
            End If
            If Not oComando.CreateCmdSp("spRep_InventarioDiaIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                Set oComando = Nothing
                Exit Sub
            End If
            oComando.CreateParameter "@vch_area", adVarChar, adParamInput, 10, IIf(optarea(0).value, "C", "S")
            oComando.CreateParameter "@vch_familia", adVarChar, adParamInput, 20, CboFamilia.BoundText
            oComando.CreateParameter "@vch_Subfamilia", adVarChar, adParamInput, 20, CboSubFamilia.BoundText
            oComando.CreateParameter "@vch_Articulo", adVarChar, adParamInput, 20, TxtCodigo.Text
            oComando.CreateParameter "@vch_Visualiza", adVarChar, adParamInput, 20, IIf(optVisualizacion(0).value, "C", "M")
            oComando.CreateParameter "@TipoStk", adVarChar, adParamInput, 3, sEvaluaFiltro
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
                    Isql = "INSERT INTO " & sNombre & "(Tienda, Codigo, Descripcion, Monto, Unidad) " & _
                           "VALUES('" & Tienda & "','" & !Codigo & "', '" & !Descripcion & "','" & !Monto & "','" & !unidad & "')"
                    Cn.Execute Isql
                    RsDatos.MoveNext
                Wend
            End With
        End If
sigue:
    Next w

    Isql = "SELECT Codigo, Descripcion, Unidad, Monto, Tienda FROM " & sNombre
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
End Sub

Private Sub optForma_Click(Index As Integer)
    Select Case Index
           Case Is = 0
                ChkFiltro(0).value = 1
                ChkFiltro(1).value = 1
                ChkFiltro(2).value = 1
                ChkFiltro(0).Enabled = True
                ChkFiltro(1).Enabled = True
                ChkFiltro(2).Enabled = True
           Case Is = 1
                ChkFiltro(0).value = 1
                ChkFiltro(1).value = 1
                ChkFiltro(2).value = 1
                ChkFiltro(0).Enabled = False
                ChkFiltro(1).Enabled = False
                ChkFiltro(2).Enabled = False
    End Select
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

Private Sub Form_Unload(Cancel As Integer)
    If sNombre <> "" Then Cn.Execute "Drop Table " & sNombre
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
