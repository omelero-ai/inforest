VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepSolicitadoVsEntregado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insumos Solicitados Vs Entregado"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmRepSolicitadoVsEntregado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   5955
      Picture         =   "frmRepSolicitadoVsEntregado.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4605
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
      Left            =   4515
      Picture         =   "frmRepSolicitadoVsEntregado.frx":0534
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4605
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
      Left            =   3075
      Picture         =   "frmRepSolicitadoVsEntregado.frx":05FC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4605
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
      Left            =   1635
      Picture         =   "frmRepSolicitadoVsEntregado.frx":0E1E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4605
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
      Left            =   195
      Picture         =   "frmRepSolicitadoVsEntregado.frx":1350
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4605
      Width           =   1455
   End
   Begin VB.Frame Frame1 
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
      Height          =   4455
      Left            =   60
      TabIndex        =   19
      Top             =   30
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Formato de Reporte"
         ForeColor       =   &H00800000&
         Height          =   705
         Left            =   840
         TabIndex        =   28
         Top             =   3630
         Width           =   6090
         Begin VB.OptionButton OptFormato 
            Caption         =   "Resumido"
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   12
            Top             =   345
            Value           =   -1  'True
            Width           =   1350
         End
         Begin VB.OptionButton OptFormato 
            Caption         =   "Detallado"
            Height          =   195
            Index           =   1
            Left            =   2430
            TabIndex        =   13
            Top             =   345
            Width           =   1290
         End
      End
      Begin VB.CheckBox ChkArticulo 
         Caption         =   "Todos"
         Height          =   210
         Left            =   6510
         TabIndex        =   9
         Top             =   2325
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
         Left            =   2505
         TabIndex        =   8
         Top             =   2280
         Width           =   3930
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
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2275
         Width           =   885
      End
      Begin VB.CheckBox ChkFamilia 
         Caption         =   "Todas"
         Height          =   195
         Left            =   6510
         TabIndex        =   5
         Top             =   1440
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox ChkSubFamilia 
         Caption         =   "Todas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6510
         TabIndex        =   7
         Top             =   1875
         Value           =   1  'Checked
         Width           =   795
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   5925
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   1620
         TabIndex        =   2
         Top             =   925
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   51511297
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1620
         TabIndex        =   0
         Top             =   465
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   51511297
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3240
         TabIndex        =   1
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
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
         CustomFormat    =   "HH:mm"
         Format          =   51511299
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3240
         TabIndex        =   3
         Top             =   930
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
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
         CustomFormat    =   "HH:mm"
         Format          =   51511299
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo CboFamilia 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   1385
         Width           =   4845
         _ExtentX        =   8546
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
         Left            =   1620
         TabIndex        =   6
         Top             =   1830
         Width           =   4845
         _ExtentX        =   8546
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   2735
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   3180
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
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
         Caption         =   "Local:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   27
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   26
         Top             =   2325
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Almacen (Local Central):"
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   285
         TabIndex        =   24
         Top             =   2670
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1035
         TabIndex        =   23
         Top             =   1395
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
         Left            =   705
         TabIndex        =   22
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   21
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   20
         Top             =   525
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmRepSolicitadoVsEntregado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSubFamilia As Recordset
Dim RsFamilia As Recordset
Dim RsArea As Recordset
Dim RsLocal As Recordset
Dim rsServer As Recordset
Dim RsDatos As Recordset
Dim rsReporte As Recordset

Dim sNombre As String
Dim sTituloG As String

Dim CrtReporteDet As New CrtInsSolicitadosVsEntregadoDetallado
Dim CrtReporteRes As New CrtInsSolicitadosVsEntregadoResumido

Sub CreaTemporal()
    sNombre = dbTemporal(sUsuario, 11, "Familia", "nVarChar(100)", _
                                       "SubFamilia", "nVarChar(100)", _
                                       "Codigo", "nVarChar(20)", _
                                       "Articulo", "nVarChar(200)", _
                                       "UMedida", "nVarChar(20)", _
                                       "Fecha", "nVarChar(20)", _
                                       "Requerimiento", "nVarChar(20)", _
                                       "NotaSalida", "nVarChar(20)", _
                                       "Pedido", "Float", _
                                       "Atendido", "Float", _
                                       "Entregado", "Float")
End Sub

Public Sub GeneraDetalle()
    Dim oComando As clsComando
    Dim w, cl
    Dim conectado As Boolean
    Dim sector As String, llocal As String, Localidad As String, sOrden As String
    
    Dim sServer As String, sBddAlm As String
    Dim CadConeccion As Connection
    Dim nCantRecepcion As Double
    Dim xx As String
    
    Set oComando = New clsComando

    If sNombre <> "" Then Cn.Execute "Delete From " & sNombre
        
    'Trae el ip y la base de datos de almacen
    Set rsServer = Lib.OpenRecordset("Select IpAlm, BdAlm From vLocalidades Where Codigo = '" & cboLocal.BoundText & "'", CnAlmacen)
    If rsServer.RecordCount > 0 Then
        sServer = rsServer!IpAlm
        sBddAlm = rsServer!bdalm
    End If
    
    If Not oComando.CreateCmdSp("spRep_RequerimientosPorAreaIntegrado", CnAlmacen) Then
        Set oComando = Nothing
        Exit Sub
    End If
    
    oComando.CreateParameter "@Area", adVarChar, adParamInput, 10, cboArea.BoundText
    oComando.CreateParameter "@FechaIni", adVarChar, adParamInput, 20, Format(dtpFecIni.value, "yyyyMMdd") & " " & Format(dtpHorIni.value, "HH:mm")
    oComando.CreateParameter "@FechaFin", adVarChar, adParamInput, 20, Format(dtpFecFin.value, "yyyyMMdd") & " " & Format(dtpHorFin.value, "HH:mm")
    oComando.CreateParameter "@Familia", adVarChar, adParamInput, 10, CboFamilia.BoundText
    oComando.CreateParameter "@Subfamilia", adVarChar, adParamInput, 10, CboSubFamilia.BoundText
    oComando.CreateParameter "@Articulo", adVarChar, adParamInput, 20, TxtCodigo.Text
    If Not oComando.GetParamOK Then
        Set oComando = Nothing
        Exit Sub
    End If
    Set RsDatos = oComando.GetSP()
    Set CadConeccion = devuelveConexion(sServer, sBddAlm)
    With RsDatos
        While Not .EOF
            nCantRecepcion = Calcular("Select Sum(nCantidad) As Codigo From dDocumento Where Substring(tDocumento, 3, 4) + Right(tdocumento, 6)= '" & !tSubSalida & "' And tCodigoProducto = '" & !tCodigoProducto & "' And tEstadoDocumento NOT IN ('01', '04')", CadConeccion)
            Isql = "INSERT INTO " & sNombre & "(Familia, SubFamilia, Codigo, Articulo, UMedida, Fecha, Requerimiento, NotaSalida, Pedido, Atendido, Entregado) VALUES('" & UCase(!familia) & "','" & UCase(!subFamilia) & "','" & !tCodigoProducto & "','" & UCase(!tDetallado) & "','" & UCase(!UnidadEntrada) & "','" & Format(!fRegistro, "dd/MM/yyyy") & "','" & !tRequerimiento & "','" & !tSubSalida & "'," & !nCantidadPedida & "," & !nAtendido & "," & nCantRecepcion & ")"
            Cn.Execute Isql
            RsDatos.MoveNext
        Wend
    End With
    
    If OptFormato(0).value Then
        Isql = "SELECT Familia, SubFamilia, Codigo, Articulo, UMedida, SUM(Pedido) AS Pedido, SUM(Atendido) AS Atendido, SUM(Entregado) AS Entregado FROM " & sNombre & " GROUP BY Familia, SubFamilia, Codigo, Articulo, UMedida"
    Else
        Isql = "SELECT Familia, SubFamilia, Codigo, Articulo, UMedida, Fecha, Requerimiento, NotaSalida, Pedido, Atendido, Entregado FROM " & sNombre
    End If
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
End Sub

Sub LlenaCombo()
    If Not CnAlmacen Is Nothing Then
        Isql = "Select Codigo, Descripcion from vFamilia where lActivo=1 Order By 2"
        Llena_Combo Isql, CboFamilia, RsFamilia, "Codigo", "Descripcion"
        
        Isql = "Select Codigo, Descripcion from vArea where lActivo=1 Order By 2"
        Llena_Combo Isql, cboArea, RsArea, "Codigo", "Descripcion"
        
        Isql = "Select Codigo, Descripcion from vLocalidades where lActivo=1 Order By 2"
        Llena_Combo Isql, cboLocal, RsLocal, "Codigo", "Descripcion"
    End If
End Sub

Private Sub CboFamilia_Change()
    CboSubFamilia.BoundText = ""
End Sub

Private Sub CboFamilia_LostFocus()
    Isql = "Select Codigo, Descripcion From vSubFamilia Where lActivo=1 And tFamilia='" & CboFamilia.BoundText & "'"
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
    Else
        ChkSubFamilia.Enabled = True
        CboFamilia.Enabled = True
        CboFamilia.BackColor = &HF0EADB
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

Private Sub cmdOpcion_Click(Index As Integer)
    If Index <> 3 Then
        sTituloG = ""
        
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
        If OptFormato(0).value Then
            sTituloG = sTituloG & " / Formato: Resumido"
        Else
            sTituloG = sTituloG & " / Formato: Detallado"
        End If
        If Len(Trim(cboArea.BoundText)) = 0 Then
            MsgBox "Seleccione el SubAlmacen!!!", vbExclamation, sMensaje
            cboArea.SetFocus
            Exit Sub
        End If
        If Len(Trim(cboLocal.BoundText)) = 0 Then
            MsgBox "Seleccione el Local!!!", vbExclamation, sMensaje
            cboLocal.SetFocus
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass
        GeneraDetalle
        
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
            If OptFormato(0).value Then
                CrtReporteRes.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteRes.Database.SetDataSource rsReporte
                CrtReporteRes.TxtRSocial.SetText sRazonComercial
                CrtReporteRes.txtSubTitulo.SetText sTituloG
                CrtReporteRes.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteRes.PaperOrientation = crPortrait
                CrtReporteRes.DiscardSavedData
                frmEmite.CRViewer.ReportSource = CrtReporteRes
            Else
                CrtReporteDet.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRSocial.SetText sRazonComercial
                CrtReporteDet.txtSubTitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteDet.PaperOrientation = crPortrait
                CrtReporteDet.DiscardSavedData
                frmEmite.CRViewer.ReportSource = CrtReporteDet
            End If
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
            
        Case Is = 1 'Impresora
            If OptFormato(0).value Then
                CrtReporteRes.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteRes.Database.SetDataSource rsReporte
                CrtReporteRes.TxtRSocial.SetText sRazonComercial
                CrtReporteRes.txtSubTitulo.SetText sTituloG
                CrtReporteRes.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteRes.PaperOrientation = crPortrait
                CrtReporteRes.DiscardSavedData
                CrtReporteRes.PrintOut
            Else
                CrtReporteDet.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRSocial.SetText sRazonComercial
                CrtReporteDet.txtSubTitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteDet.PaperOrientation = crPortrait
                CrtReporteDet.DiscardSavedData
                CrtReporteDet.PrintOut
            End If
            
        Case Is = 2 'Exportar
            If OptFormato(0).value Then
                CrtReporteRes.ExportOptions.FormatType = crEFTExcel50
                CrtReporteRes.ExportOptions.DestinationType = crEDTDiskFile
                CrtReporteRes.Database.SetDataSource rsReporte
                CrtReporteRes.TxtRSocial.SetText sRazonComercial
                CrtReporteRes.txtSubTitulo.SetText sTituloG
                CrtReporteRes.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteRes.PaperOrientation = crLandscape
                CrtReporteRes.DiscardSavedData
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                CrtReporteRes.ExportOptions.DiskFileName = cmdSave.FileName
                If Trim(CrtReporteRes.ExportOptions.DiskFileName) <> "" Then CrtReporteRes.Export False
            Else
                CrtReporteDet.ExportOptions.FormatType = crEFTExcel50
                CrtReporteDet.ExportOptions.DestinationType = crEDTDiskFile
                CrtReporteDet.Database.SetDataSource rsReporte
                CrtReporteDet.TxtRSocial.SetText sRazonComercial
                CrtReporteDet.txtSubTitulo.SetText sTituloG
                CrtReporteDet.txtTitulo.SetText "SubAlmacen: " & UCase(cboArea.Text) & " Vs Local: " & UCase(cboLocal.Text)
                CrtReporteDet.PaperOrientation = crLandscape
                CrtReporteDet.DiscardSavedData
                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                cmdSave.ShowSave
                CrtReporteDet.ExportOptions.DiskFileName = cmdSave.FileName
                If Trim(CrtReporteDet.ExportOptions.DiskFileName) <> "" Then CrtReporteDet.Export False
            End If
            
        Case Is = 3 'Salir
            Unload Me
    End Select
    Screen.MousePointer = vbDefault
    Set rsReporte = Nothing
End Sub

Private Sub Form_Load()
    Centrar Me
    
    LlenaCombo
    
    CreaTemporal
    
    dtpFecIni.value = "01/" & str(Month(Date)) & "/" & str(Year(Date))
    dtpFecFin.value = Date
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

